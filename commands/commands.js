Office.onReady();

// 1. 發送攔截 
function validateSend(event) {
    // 讀取這封信的自訂屬性 'isVerified'
    Office.context.mailbox.item.loadCustomPropertiesAsync(async (result) => {
        const props = result.value;
        const isVerified = props.get("isVerified");
        const verifiedStateJson = props.get("verifiedState");

        if (isVerified === true && verifiedStateJson) {
            // 讀取目前的狀態進行比對
            try {
                const currentState = await getCurrentState();
                const savedState = JSON.parse(verifiedStateJson);

                // 比對各項欄位
                const isMatch = (
                    currentState.recipients === savedState.recipients &&
                    currentState.attachments === savedState.attachments &&
                    currentState.subject === savedState.subject &&
                    currentState.bodyFingerprint === savedState.bodyFingerprint
                );

                if (isMatch) {
                    // 驗證通過且內容未更動 -> 放行
                    event.completed({ allowEvent: true });
                } else {
                    // 內容已更動 -> 阻擋並重設狀態
                    props.set("isVerified", false);
                    props.saveAsync(() => {
                        event.completed({
                            allowEvent: false,
                            errorMessage: "Email content or recipients have changed since verification. Please open the 'K-Antimisdeliv' checklist to re-verify before sending."
                        });
                    });
                }
            } catch (e) {
                // 發生錯誤，阻擋發送
                event.completed({
                    allowEvent: false,
                    errorMessage: "Verification error. Please re-open the checklist and verify again."
                });
            }
        } else {
            // 未驗證 -> 阻擋
            event.completed({
                allowEvent: false,
                errorMessage: "Please click the 'K-Antimisdeliv' button above to confirm recipients and attachments before sending."
            });
        }
    });
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

    const getEmails = (arr) => (arr || []).map(p => p.emailAddress.toLowerCase()).sort().join(";");
    const getAtts = (arr) => (arr || []).map(a => a.name + a.size).sort().join(";");

    const bodyFingerprint = body ? `${body.length}_${body.substring(0, 50)}` : "empty";

    return {
        recipients: `to:${getEmails(to)}|cc:${getEmails(cc)}|bcc:${getEmails(bcc)}`,
        attachments: getAtts(attachments),
        subject: subject || "",
        bodyFingerprint: bodyFingerprint
    };
}

if (typeof g === 'undefined') var g = window;
g.validateSend = validateSend;