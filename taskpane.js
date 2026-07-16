/* global Office, document */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        try {
            loadItemData();
            document.getElementById("btnVerify").onclick = markAsVerified;

            // 1. 註冊事件監聽 (收件者與附件可即時同步)
            const item = Office.context.mailbox.item;
            item.addHandlerAsync(Office.EventType.RecipientsChanged, onMessageChanged);
            item.addHandlerAsync(Office.EventType.AttachmentsChanged, onMessageChanged);

            // 2. 設定輪詢 (每 3 秒同步主旨與內容)
            setInterval(pollForChanges, 3000);

        } catch (e) {
            logError("Init Error: " + e.message);
        }
    }
});

let lastSeenState = null;

async function pollForChanges() {
    const currentState = await getCurrentState();
    if (!lastSeenState) {
        lastSeenState = currentState;
        return;
    }

    // 比對是否有任何變動 (主旨、內容、收件者、附件)
    const isDifferent = (
        currentState.recipients !== lastSeenState.recipients ||
        currentState.attachments !== lastSeenState.attachments ||
        currentState.subject !== lastSeenState.subject ||
        currentState.bodyFingerprint !== lastSeenState.bodyFingerprint
    );

    if (isDifferent) {
        console.log("Detected change via polling, refreshing...");
        lastSeenState = currentState;
        onMessageChanged();
    }
}

function onMessageChanged() {
    // When recipients or attachments change, reset verification and reload
    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        props.set("isVerified", false);
        props.saveAsync(() => {
            // Re-show verification area and reload data
            document.getElementById("btn-area").style.display = "block";
            document.getElementById("status-msg").style.display = "none";
            loadItemData();
        });
    });
}

function logError(msg) {
    const el = document.getElementById("error-log");
    el.style.display = "block";
    el.innerText += "❌ " + msg + "\n";
    console.error(msg);
    reportErrorToApi(msg, "Sidebar/Taskpane");
}

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

function getDomain(email) {
    if (!email || typeof email !== 'string') return "unknown";
    if (!email.includes("@")) return "unknown";
    return email.split("@")[1].toLowerCase().trim();
}

function loadItemData() {
    const item = Office.context.mailbox.item;

    if (!item) {
        logError("Unable to read mail object (Item is null)");
        return;
    }

    const safeGet = (apiCall) => new Promise(resolve => {
        try {
            apiCall(result => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    console.warn("API Failed:", result.error);
                    resolve(null);
                }
            });
        } catch (e) {
            console.error("API Call Error:", e);
            resolve(null);
        }
    });

    // 重新加入附件讀取
    Promise.all([
        safeGet(cb => item.from.getAsync(cb)),
        safeGet(cb => item.to.getAsync(cb)),
        safeGet(cb => item.cc.getAsync(cb)),
        safeGet(cb => item.bcc.getAsync(cb)),
        safeGet(cb => item.getAttachmentsAsync(cb)),
        safeGet(cb => item.subject.getAsync(cb)),
        safeGet(cb => item.body.getAsync(Office.CoercionType.Html, cb))
    ]).then(([from, to, cc, bcc, attachments, subject, htmlBody]) => {

        to = to || [];
        cc = cc || [];
        bcc = bcc || [];
        attachments = attachments || [];

        // Render Subject
        document.getElementById("subject-container").innerText = subject || "(No Subject)";

        // 解析 HTML 內容
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlBody || "", 'text/html');

        // 移除所有的 style 與 script 標籤，避免其內容被 innerText 當成純文字讀取出來
        doc.querySelectorAll('style, script').forEach(el => el.remove());

        // 1. 偵測內文中的連結 (可能是雲端附件)
        const detectedLinks = [];
        const fileExtensions = ['svg', 'pdf', 'docx', 'xlsx', 'pptx', 'zip', 'rar', '7z', 'png', 'jpg', 'jpeg', 'gif'];

        doc.querySelectorAll('a').forEach(a => {
            const text = a.innerText.trim();
            const href = a.getAttribute('href') || "";
            const isFile = fileExtensions.some(ext => text.toLowerCase().endsWith('.' + ext) || href.toLowerCase().includes('.' + ext));

            if (isFile && text) {
                detectedLinks.push({
                    name: text,
                    size: 0,
                    id: href,
                    isDetected: true
                });
            }
        });

        // 2. 取得乾淨文字內容 (用於內容檢查區)
        // 移除這些被偵測為檔案的節點
        doc.querySelectorAll('a').forEach(a => {
            const text = a.innerText.trim();
            if (fileExtensions.some(ext => text.toLowerCase().endsWith('.' + ext))) {
                a.remove();
            }
        });

        const cleanText = doc.body.innerText.trim();
        document.getElementById("body-container").innerText = cleanText || "(No Content)";

        // 3. 整合附件清單
        const finalAttachments = [...attachments, ...detectedLinks];

        const senderEmail = (from && from.emailAddress) ? from.emailAddress : "";
        const senderDomain = getDomain(senderEmail);

        renderSender("from-container", from);
        renderGroupedList("to-list", to, senderDomain);
        renderGroupedList("cc-list", cc, senderDomain);
        renderGroupedList("bcc-list", bcc, senderDomain);

        // 執行附件渲染 (使用整合後的清單)
        renderAttachments("attachment-list", finalAttachments);

        checkAllChecked();

        // 更新最後看到的狀態快照，避免輪詢立即觸發
        getCurrentState().then(state => {
            lastSeenState = state;
        });

    }).catch(err => {
        logError("Load Data Error: " + err.message);
    });
}

function renderSender(containerId, data) {
    const container = document.getElementById(containerId);
    if (!data) {
        container.innerHTML = "<div class='empty-msg'>Sender info loading or not set</div>";
        return;
    }
    container.innerHTML = `
        <div class="safe-icon">👤</div>
        <div class="item-content">
            <div class="name">${data.displayName || data.emailAddress}</div>
            <div class="email">${data.emailAddress}</div>
        </div>
    `;
}

function renderGroupedList(containerId, dataArray, senderDomain) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    if (!dataArray || dataArray.length === 0) {
        container.innerHTML = "<div class='empty-msg'>(None)</div>";
        return;
    }

    const groups = {};
    dataArray.forEach(p => {
        const domain = getDomain(p.emailAddress);
        if (!groups[domain]) groups[domain] = [];
        groups[domain].push(p);
    });

    // 排序：External 排前面
    const sortedDomains = Object.keys(groups).sort((a, b) => {
        const aIsExt = a !== senderDomain;
        const bIsExt = b !== senderDomain;
        return bIsExt - aIsExt;
    });

    sortedDomains.forEach(domain => {
        const isExternal = domain !== senderDomain;
        const recipients = groups[domain];

        const groupDiv = document.createElement("div");
        groupDiv.className = "domain-group";

        const headerDiv = document.createElement("div");
        headerDiv.className = "domain-header";

        const tagHtml = isExternal
            ? `<span class="tag external">External</span>`
            : `<span class="tag internal">Internal</span>`;

        // 將勾選框移至 Header
        const checkedState = isExternal ? "" : "checked";
        headerDiv.innerHTML = `
            <div style="display: flex; align-items: center;">
                <input type='checkbox' class='verify-check' ${checkedState} onchange='checkAllChecked()'>
                <span>@${domain}</span>
            </div>
            ${tagHtml}
        `;
        groupDiv.appendChild(headerDiv);

        recipients.forEach((p, i) => {
            const rowDiv = document.createElement("div");
            rowDiv.className = "item-row";

            // 移除個別勾選框，並依賴 CSS 的 padding 縮進
            rowDiv.innerHTML = `
                <div class="item-content">
                    <div class="name">${p.displayName || p.emailAddress}</div>
                    <div class="email">${p.emailAddress}</div>
                </div>
            `;
            groupDiv.appendChild(rowDiv);
        });

        container.appendChild(groupDiv);
    });
}

// 移除 renderAttachments 函式

window.checkAllChecked = function () {
    const allCheckboxes = document.querySelectorAll(".verify-check");
    let pass = true;

    if (allCheckboxes.length === 0) {
        pass = true;
    } else {
        allCheckboxes.forEach(c => {
            if (!c.checked) pass = false;
        });
    }

    if (pass) enableButton();
    else disableButton();
};

function enableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = false;
    btn.classList.add("active");
    btn.innerText = "Verify information";
}

function disableButton() {
    const btn = document.getElementById("btnVerify");
    btn.disabled = true;
    btn.classList.remove("active");

    const all = document.querySelectorAll(".verify-check");
    let uncheckCount = 0;
    all.forEach(c => { if (!c.checked) uncheckCount++; });

    btn.innerText = uncheckCount > 0 ? `${uncheckCount} items left to verify` : "Please check all items...";
}

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

    // Simple fingerprint for body to detect changes without storing huge strings
    const bodyFingerprint = body ? `${body.length}_${body.substring(0, 50)}` : "empty";

    return {
        recipients: `to:${getEmails(to)}|cc:${getEmails(cc)}|bcc:${getEmails(bcc)}`,
        attachments: getAtts(attachments),
        subject: subject || "",
        bodyFingerprint: bodyFingerprint
    };
}

function hashCode(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = (hash << 5) - hash + str.charCodeAt(i);
        hash |= 0; // Convert to 32bit integer
    }
    return hash.toString(36); // 轉成 36 進位字串更短
}

async function markAsVerified() {
    const state = await getCurrentState();

    // 由於 CustomProperties 單個屬性有 28KB 大小限制，
    // 我們將特徵組合字串進行 Hash 壓縮，只儲存一個極短的指紋字串（例如 "9u2p1a"），不佔空間且非常安全。
    const rawStateString = `${state.recipients}_${state.attachments}_${state.subject}_${state.bodyFingerprint}`;
    const hashedState = hashCode(rawStateString);

    Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
        const props = result.value;
        props.set("isVerified", true);
        props.set("verifiedState", hashedState);

        props.saveAsync((saveResult) => {
            if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("btn-area").style.display = "none";
                document.getElementById("status-msg").style.display = "block";
            } else {
                logError("Save failed: " + saveResult.error.message);
            }
        });
    });
}

function getFileIcon(filename) {
    if (!filename) return "assets/ic_file.svg";
    const ext = filename.split('.').pop().toLowerCase();

    // File Type Mapping (SVG)
    switch (ext) {
        case 'ai': return "assets/ic_ai.svg";
        case 'csv': case 'xls': case 'xlsx': return "assets/ic_csv.svg";
        case 'pdf': return "assets/ic_pdf.svg";
        case 'txt': case 'log': case 'md': case 'rtf': return "assets/ic_txt.svg";
        case 'mp3': case 'wav': case 'ogg': return "assets/ic_audio.svg";
        case 'exe': case 'msi': return "assets/ic_exe.svg";
        case 'ppt': case 'pptx': return "assets/ic_ppt.svg";
        case 'mp4': case 'mov': case 'avi': case 'mkv': return "assets/ic_video.svg";
        case 'js': case 'html': case 'css': case 'json': case 'xml': case 'ts': return "assets/ic_code.svg";
        case 'fig': return "assets/ic_fig.svg";
        case 'jpg': case 'jpeg': case 'png': case 'gif': case 'bmp': case 'svg': return "assets/ic_img.svg";
        case 'rar': return "assets/ic_rar.svg";
        case 'zip': case '7z': case 'tar': case 'gz': return "assets/ic_zip.svg";
        default: return "assets/ic_file.svg";
    }
}

function renderAttachments(containerId, attachments) {
    const container = document.getElementById(containerId);
    container.innerHTML = "";

    if (!attachments || attachments.length === 0) {
        container.innerHTML = "<div class='empty-msg'>(No Attachments)</div>";
        return;
    }

    // 建立附件區塊的 Header (包含全選勾選框)
    const headerDiv = document.createElement("div");
    headerDiv.className = "domain-header"; // 延用 domain-header 的樣式
    headerDiv.innerHTML = `
        <div style="display: flex; align-items: center;">
            <input type='checkbox' class='verify-check' onchange='checkAllChecked()'>
            <span>Check All Attachments</span>
        </div>
        <span class="tag internal" style="background:#e0e0e0; color:#666;">${attachments.length} files</span>
    `;
    container.appendChild(headerDiv);

    attachments.forEach((att, i) => {
        const rowDiv = document.createElement("div");
        rowDiv.className = "item-row";
        rowDiv.style.paddingLeft = "30px"; // 讓內容縮排，對齊 Header 的文字

        const iconPath = getFileIcon(att.name);
        const typeTag = att.isDetected ? `<span class="tag internal" style="font-size:8px; margin-left:5px;">Content</span>` : "";

        // 移除個別勾選框，改為單純顯示資訊
        rowDiv.innerHTML = `
            <div class="item-content">
                <div class="name">
                    ${att.name} ${typeTag}
                </div>
                <div class="email">${att.isDetected ? "Link to OneDrive" : (att.size / 1024).toFixed(1) + " KB"}</div>
            </div>
        `;
        container.appendChild(rowDiv);
    });
}
