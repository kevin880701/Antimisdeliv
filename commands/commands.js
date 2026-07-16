Office.onReady();

function hashCode(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = (hash << 5) - hash + str.charCodeAt(i);
        hash |= 0;
    }
    return hash.toString(36);
}

let pendingLogPromise = null;

// 封裝 event.completed，確保在關閉背景進程前，API 請求已經順利送達
async function completeEvent(event, options) {
    if (pendingLogPromise) {
        try {
            // 等待最長 2 秒以確保日誌送出
            await Promise.race([
                pendingLogPromise,
                new Promise(resolve => setTimeout(resolve, 2000))
            ]);
        } catch (e) {
            console.error("Failed to wait for log:", e);
        }
    }
    event.completed(options);
}

// 1. 發送攔截 
function validateSend(event) {
    pendingLogPromise = null; // 重設 Promise 狀態
    try {
        // 讀取這封信的自訂屬性 'isVerified'
        Office.context.mailbox.item.loadCustomPropertiesAsync(async (result) => {
            try {
                if (result.status === Office.AsyncResultStatus.Failed || !result.value) {
                    const errMsg = "Failed to load verification properties. Please open the 'Antimisdeliv' checklist to verify again.";
                    pendingLogPromise = reportErrorToApi(errMsg + " (Status: " + result.status + ")", "Send/OnMessageSend");
                    completeEvent(event, {
                        allowEvent: false,
                        errorMessage: errMsg
                    });
                    return;
                }

                const props = result.value;
                const isVerified = props.get("isVerified");
                const verifiedStateHash = props.get("verifiedState");

                if (isVerified === true && verifiedStateHash) {
                    try {
                        const currentState = await getCurrentState();
                        const rawStateString = `${currentState.recipients}_${currentState.attachments}_${currentState.subject}_${currentState.bodyFingerprint}`;
                        const currentHash = hashCode(rawStateString);

                        // 比對雜湊值是否相符
                        if (currentHash === verifiedStateHash) {
                            // 驗證通過且內容未更動 -> 放行
                            completeEvent(event, { allowEvent: true });
                        } else {
                            // 內容已更動 -> 阻擋並重設狀態
                            props.set("isVerified", false);
                            props.saveAsync(async () => {
                                const errMsg = "Email content or recipients have changed since verification. Please open the 'Antimisdeliv' checklist to re-verify before sending.";
                                pendingLogPromise = reportErrorToApi(errMsg, "Send/HashMismatch");
                                completeEvent(event, {
                                    allowEvent: false,
                                    errorMessage: errMsg
                                });
                            });
                        }
                    } catch (e) {
                        // 發生錯誤，阻擋發送
                        const errMsg = "Verification error (" + e.toString() + "). Please re-open the checklist and verify again.";
                        pendingLogPromise = reportErrorToApi("getCurrentState/hashCode error: " + e.toString(), "Send/VerificationError");
                        completeEvent(event, {
                            allowEvent: false,
                            errorMessage: errMsg
                        });
                    }
                } else {
                    // 未驗證 -> 阻擋
                    completeEvent(event, {
                        allowEvent: false,
                        errorMessage: "Please click the 'Antimisdeliv' button above to confirm recipients and attachments before sending."
                    });
                }
            } catch (e) {
                pendingLogPromise = reportErrorToApi("Inner validateSend exception: " + e.toString(), "Send/InnerException");
                pendingLogPromise.then(() => {
                    completeEvent(event, {
                        allowEvent: false,
                        errorMessage: "Inner error: " + e.toString()
                    });
                });
            }
        });
    } catch (e) {
        pendingLogPromise = reportErrorToApi("Outer validateSend exception: " + e.toString(), "Send/OuterException");
        pendingLogPromise.then(() => {
            completeEvent(event, {
                allowEvent: false,
                errorMessage: "Outer error: " + e.toString()
            });
        });
    }
}

// 傳送錯誤日誌到 API
async function reportErrorToApi(errorMessage, source) {
    try {
        const userEmail = (Office && Office.context && Office.context.mailbox && Office.context.mailbox.userProfile)
            ? Office.context.mailbox.userProfile.emailAddress
            : "unknown_user";

        // 💡 步驟 1：第一時間立刻發送「已觸發」通知，避免後續 API 卡死導致完全沒收到日誌
        fetch("https://startingacademy.ai-edm.com/api/v1/test/antimisdeliv", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                message: `[${source}] 使用者: ${userEmail} - 動作觸發: ${errorMessage} (開始收集詳細元資料...)`,
                status: "info",
                code: 200
            })
        }).catch(() => {});

        // 💡 步驟 2：收集詳細的信件狀態
        let mailMetadata = {};
        try {
            mailMetadata = await getMailMetadataForLogging();
        } catch (metaErr) {
            mailMetadata = { error: "Failed to collect metadata: " + metaErr.toString() };
        }

        // 💡 步驟 3：發送包含詳細信件資料的完整 Log
        fetch("https://startingacademy.ai-edm.com/api/v1/test/antimisdeliv", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                message: `[${source}] 使用者: ${userEmail}\n狀態: ${errorMessage}\n\n【信件除錯狀態】:\n${JSON.stringify(mailMetadata, null, 2)}`,
                status: "error",
                code: 500,
                mailMetadata: mailMetadata
            })
        }).catch(err => console.error("Error reporting to API failed:", err));
    } catch (e) {
        console.error("reportErrorToApi failed:", e);
    }
}

// 💡 安全收集信件所有屬性的狀況，加強防護：若 1000ms 未回傳則自動判定 Timeout 續行，不影響日誌發送
async function getMailMetadataForLogging() {
    const item = Office.context.mailbox.item;
    if (!item) return { error: "mailbox.item is null" };

    const safeGetLog = (apiCall) => new Promise(resolve => {
        let resolved = false;
        
        // 💡 1秒強制逾時限制，避免 Promise 永久懸空
        const timer = setTimeout(() => {
            if (!resolved) {
                resolved = true;
                resolve({ status: "timeout", error: "API call timed out after 1000ms" });
            }
        }, 1000);

        try {
            apiCall(result => {
                if (!resolved) {
                    resolved = true;
                    clearTimeout(timer);
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve({ status: "success", value: result.value });
                    } else {
                        resolve({ status: "failed", error: result.error ? result.error.message : "unknown error" });
                    }
                }
            });
        } catch (e) {
            if (!resolved) {
                resolved = true;
                clearTimeout(timer);
                resolve({ status: "exception", error: e.toString() });
            }
        }
    });

    const [to, cc, bcc, subject, attachments, body] = await Promise.all([
        safeGetLog(cb => item.to.getAsync(cb)),
        safeGetLog(cb => item.cc.getAsync(cb)),
        safeGetLog(cb => item.bcc.getAsync(cb)),
        safeGetLog(cb => item.subject.getAsync(cb)),
        safeGetLog(cb => item.getAttachmentsAsync(cb)),
        safeGetLog(cb => item.body.getAsync(Office.CoercionType.Text, cb))
    ]);

    return {
        to: to,
        cc: cc,
        bcc: bcc,
        subject: subject,
        attachments: attachments,
        body: body
    };
}

// 取得目前郵件狀態
async function getCurrentState() {
    const item = Office.context.mailbox.item;
    const safeGet = (apiCall) => new Promise(resolve => {
        try {
            apiCall(result => resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null));
        } catch (e) {
            resolve(null);
        }
    });

    const [to, cc, bcc, attachments, subject, body] = await Promise.all([
        safeGet(cb => item.to.getAsync(cb)),
        safeGet(cb => item.cc.getAsync(cb)),
        safeGet(cb => item.bcc.getAsync(cb)),
        safeGet(cb => item.getAttachmentsAsync(cb)),
        safeGet(cb => item.subject.getAsync(cb)),
        safeGet(cb => item.body.getAsync(Office.CoercionType.Text, cb))
    ]);

    const getEmails = (arr) => (arr || []).filter(p => p && p.emailAddress).map(p => p.emailAddress.toLowerCase()).sort().join(";");
    const getAtts = (arr) => (arr || []).map(a => a.name + a.size).sort().join(";");

    const bodyFingerprint = body ? `${body.length}_${body.substring(0, 50)}` : "empty";

    return {
        recipients: `to:${getEmails(to)}|cc:${getEmails(cc)}|bcc:${getEmails(bcc)}`,
        attachments: getAtts(attachments),
        subject: subject || "",
        bodyFingerprint: bodyFingerprint
    };
}

const g = typeof globalThis !== 'undefined' ? globalThis : 
          typeof window !== 'undefined' ? window : 
          typeof global !== 'undefined' ? global : 
          typeof self !== 'undefined' ? self : this;

g.validateSend = validateSend;