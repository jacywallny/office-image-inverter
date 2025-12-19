Office.onReady((info) => {
    const btn = document.getElementById("runBtn");
    const desc = document.querySelector("p");
    
    // ============================================================
    // 🧠 核心分流逻辑：根据宿主环境决定按钮的功能
    // ============================================================
    
    if (info.host === Office.HostType.Word) {
        // 🟦 情况 A：在 Word 里
        // 使用“强力内核”，点击直接处理
        console.log("环境检测: Word (启用强力模式)");
        if (btn) {
            btn.innerText = "一键反色 (Word)";
            btn.onclick = runInvertInWord; // 绑定 Word 专用函数
        }
        if (desc) desc.innerText = "选中 Word 图片 -> 点击按钮";

    } else {
        // 🟧 情况 B：在 PPT (或 Excel) 里
        // 使用“剪贴板/拖拽模式”
        console.log("环境检测: PPT/其他 (启用剪贴板模式)");
        if (btn) {
            btn.innerText = "🖱️ 点我，然后按 Ctrl+V";
            btn.onclick = () => {
                updateStatus("👉 没错！请直接按下 Ctrl+V 粘贴图片");
            };
        }
        if (desc) desc.innerText = "PPT操作：复制图片 -> 点这里 -> 按 Ctrl+V";
        
        // 只有在非 Word 环境下，才监听粘贴事件
        document.addEventListener("paste", handlePaste);
    }

    // ============================================================
    // 🖱️ 全局功能：拖拽支持 (Word 和 PPT 都能用，作为备选)
    // ============================================================
    setupDragAndDrop();
});


// ==========================================
// 🟦 模式一：Word 专用强力内核 (Word.run)
// ==========================================
async function runInvertInWord() {
    updateStatus("⏳ Word模式：正在读取...");
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const pictures = selection.inlinePictures;
            pictures.load("items");
            await context.sync();

            if (pictures.items.length === 0) {
                updateStatus("❌ 未检测到图片！\n请右键图片 -> 自动换行 -> 设为【嵌入型】");
                return;
            }

            const wordPicture = pictures.items[0];
            const base64Result = wordPicture.getBase64ImageSrc();
            await context.sync();

            const base64 = base64Result.value;
            if (!base64) {
                updateStatus("❌ 无法读取图片数据");
                return;
            }

            updateStatus("🎨 正在反色...");
            const newBase64 = await invertImagePromise(base64);

            const cleanBase64 = newBase64.split(",")[1];
            wordPicture.insertInlinePictureFromBase64(cleanBase64, "Replace");
            await context.sync();
            updateStatus("✅ 成功！已反色");
        });
    } catch (error) {
        console.error(error);
        updateStatus("⚠️ Word内核报错: " + error.message);
    }
}


// ==========================================
// 🟧 模式二：剪贴板粘贴处理 (PPT专用)
// ==========================================
async function handlePaste(event) {
    event.preventDefault(); // 阻止默认粘贴
    const items = (event.clipboardData || event.originalEvent.clipboardData).items;
    let blob = null;

    for (const item of items) {
        if (item.type.indexOf("image") === 0) {
            blob = item.getAsFile();
            break;
        }
    }

    if (blob) {
        await processBlobToClipboard(blob);
    } else {
        updateStatus("❌ 粘贴板里没有图片！请先在 PPT 复制。");
    }
}


// ==========================================
// 🖱️ 拖拽功能 (通用)
// ==========================================
function setupDragAndDrop() {
    document.body.addEventListener("dragover", (e) => {
        e.preventDefault();
        document.body.style.backgroundColor = "#e6f2ff";
        updateStatus("✊ 松手即可处理");
    });

    document.body.addEventListener("dragleave", (e) => {
        e.preventDefault();
        document.body.style.backgroundColor = "";
        updateStatus("等待操作...");
    });

    document.body.addEventListener("drop", async (e) => {
        e.preventDefault();
        document.body.style.backgroundColor = "";
        const items = e.dataTransfer.items;
        let blob = null;
        for (const item of items) {
            if (item.type.indexOf("image") === 0) {
                blob = item.getAsFile();
                break;
            }
        }
        if (blob) {
            await processBlobToClipboard(blob);
        } else {
            updateStatus("❌ 拖进来的不是图片");
        }
    });
}


// ==========================================
// 🛠️ 核心算法与工具
// ==========================================

// 统一处理：Blob -> 反色 -> 剪贴板
async function processBlobToClipboard(blob) {
    try {
        updateStatus("🎨 正在反色...");
        const base64 = await blobToBase64(blob);
        const newBase64 = await invertImagePromise(base64);
        const newBlob = await base64ToBlob(newBase64);
        
        await navigator.clipboard.write([
            new ClipboardItem({ [blob.type]: newBlob })
        ]);

        updateStatus("✅ 成功！请按 Ctrl+V 粘贴");
        
        // 按钮绿色反馈
        const btn = document.getElementById("runBtn");
        if(btn) {
            const oldBg = btn.style.backgroundColor;
            const oldTxt = btn.innerText;
            btn.style.backgroundColor = "#107c10";
            btn.innerText = "完成！请粘贴";
            setTimeout(() => {
                btn.style.backgroundColor = oldBg;
                btn.innerText = oldTxt;
            }, 2000);
        }

    } catch (err) {
        console.error(err);
        updateStatus("⚠️ 处理出错: " + err.message);
    }
}

function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

async function base64ToBlob(base64) {
    const res = await fetch(base64);
    return await res.blob();
}

function invertImagePromise(base64Str) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0);
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const data = imageData.data;
            for (let i = 0; i < data.length; i += 4) {
                data[i] = 255 - data[i];
                data[i + 1] = 255 - data[i + 1];
                data[i + 2] = 255 - data[i + 2];
            }
            ctx.putImageData(imageData, 0, 0);
            resolve(canvas.toDataURL("image/png"));
        };
        img.onerror = reject;
        img.src = base64Str;
    });
}

function updateStatus(msg) {
    // 尝试寻找美化版 UI 的元素，如果没有就找简陋版的
    const el = document.getElementById("status");
    if (el) el.innerText = msg;
    
    // 如果你有美化版 UI，这里可以增加更多逻辑，比如转圈圈的显示/隐藏
    const spinner = document.getElementById("spinner");
    if (spinner) {
        spinner.style.display = (msg.includes("正在") || msg.includes("...")) ? "block" : "none";
    }
}
