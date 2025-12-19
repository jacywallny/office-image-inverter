Office.onReady((info) => {
    // 1. 初始化界面文字
    const btn = document.getElementById("runBtn");
    const desc = document.querySelector("p");

    if (btn) {
        btn.innerText = "🖱️ 点我，然后按 Ctrl+V";
        btn.onclick = () => {
            updateStatus("👉 没错！请直接按下 Ctrl+V 粘贴，或者把图片拖进来");
        };
    }
    
    if (desc) desc.innerText = "方法一：复制图片 -> 点这里 -> 按 Ctrl+V\n方法二：直接把 PPT 里的图片拖进来";
    
    // 2. 监听粘贴 (Ctrl+V)
    document.addEventListener("paste", handlePaste);

    // ==========================================
    // 🆕 新增功能：监听拖拽 (Drag & Drop)
    // ==========================================
    
    // 当文件拖进插件区域时：变色提示
    document.body.addEventListener("dragover", (e) => {
        e.preventDefault(); // 必须加这行，允许拖入
        document.body.style.backgroundColor = "#e6f2ff"; // 变成淡蓝色
        updateStatus("✊ 松手即可处理图片");
    });

    // 当文件离开插件区域时：恢复颜色
    document.body.addEventListener("dragleave", (e) => {
        e.preventDefault();
        document.body.style.backgroundColor = ""; // 恢复原色
        updateStatus("等待图片...");
    });

    // 当文件被扔下 (松手) 时：
    document.body.addEventListener("drop", async (e) => {
        e.preventDefault(); // 阻止浏览器默认打开图片的行为
        document.body.style.backgroundColor = ""; // 恢复原色
        
        updateStatus("⚡ 捕获到拖拽对象，正在分析...");

        // 获取拖拽的数据
        const items = e.dataTransfer.items;
        let blob = null;

        // 寻找是不是图片
        for (const item of items) {
            if (item.type.indexOf("image") === 0) {
                blob = item.getAsFile();
                break;
            }
        }

        if (blob) {
            // 如果是图片，直接复用我们的核心处理逻辑
            await processBlobToClipboard(blob);
        } else {
            updateStatus("❌ 拖进来的不是图片！\n请拖拽 PPT 里的图片或截图文件。");
        }
    });
});

// ==========================================
// 核心逻辑区域
// ==========================================

// 处理粘贴事件
async function handlePaste(event) {
    event.preventDefault();
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
        updateStatus("❌ 粘贴板里没有图片！");
    }
}

// 统一处理函数：拿到图片文件(Blob) -> 反色 -> 塞回剪贴板
async function processBlobToClipboard(blob) {
    try {
        updateStatus("🎨 正在进行反色计算...");

        // 1. 转 Base64
        const base64 = await blobToBase64(blob);
        
        // 2. 反色
        const newBase64 = await invertImagePromise(base64);

        // 3. 塞回剪贴板
        const newBlob = await base64ToBlob(newBase64);
        await navigator.clipboard.write([
            new ClipboardItem({ [blob.type]: newBlob })
        ]);

        // 4. 成功提示
        updateStatus("✅ 成功！新图已放入剪贴板。\n请回到 PPT 按 Ctrl+V");
        
        // 按钮变绿反馈
        const btn = document.getElementById("runBtn");
        if(btn) {
            btn.style.backgroundColor = "#107c10";
            const oldText = btn.innerText;
            btn.innerText = "完成！请粘贴";
            setTimeout(() => {
                btn.style.backgroundColor = "";
                btn.innerText = oldText;
            }, 3000);
        }

    } catch (err) {
        console.error(err);
        updateStatus("⚠️ 处理出错: " + err.message);
    }
}

// --- 辅助工具函数 (不需要动) ---

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
    const el = document.getElementById("status");
    if (el) el.innerText = msg;
}
