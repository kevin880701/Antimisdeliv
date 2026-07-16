Office.onReady();

function hashCode(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = (hash << 5) - hash + str.charCodeAt(i);
        hash |= 0;
    }
    return hash.toString(36);
}

// 1. 發送攔截 
function validateSend(event) {
    try {
        // 讀取這封信的自訂屬性 'isVerified'
        Office.context.mailbox.item.loadCustomPropertiesAsync(async (result) => {
            try {
                if (result.status === Office.AsyncResultStatus.Failed || !result.value) {
                    event.completed({
                        allowEvent: false,
                        errorMessage: "Failed to load verification properties. Please open the 'Antimisdeliv' checklist to verify again."
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
                            event.completed({ allowEvent: true });
                        } else {
                            // 內容已更動 -> 阻擋並重設狀態
                            props.set("isVerified", false);
                            props.saveAsync(() => {
                                event.completed({
                                    allowEvent: false,
                                    errorMessage: "Email content or recipients have changed since verification. Please open the 'Antimisdeliv' checklist to re-verify before sending."
                                });
                            });
                        }
                    } catch (e) {
                        // 發生錯誤，阻擋發送
                        event.completed({
                            allowEvent: false,
                            errorMessage: "Verification error (" + e.toString() + "). Please re-open the checklist and verify again."
                        });
                    }
                } else {
                    // 未驗證 -> 阻擋
                    event.completed({
                        allowEvent: false,
                        errorMessage: "Please click the 'Antimisdeliv' button above to confirm recipients and attachments before sending."
                    });
                }
            } catch (e) {
                event.completed({
                    allowEvent: false,
                    errorMessage: "Inner error: " + e.toString()
                });
            }
        });
    } catch (e) {
        event.completed({
            allowEvent: false,
            errorMessage: "Outer error: " + e.toString()
        });
    }
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