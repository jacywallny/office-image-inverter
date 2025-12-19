Office.onReady((info) => {
    // 1. 初始化：改变界面提示，告诉用户怎么用
    const btn = document.getElementById("runBtn");
    const status = document.getElementById("status");
    const title = document.querySelector("h2"); // 假设你有h2标题
    const desc = document.querySelector("p");   // 假设你有p标签说明

    if (btn) {
        // 既然不能自动读，就把按钮改成一个“状态指示器”
        btn.innerText = "🖱️ 点我，然后按 Ctrl+V";
        btn.onclick = () => {
            updateStatus("👉 没错！请直接按下 Ctrl+V 粘贴图片");
        };
    }
    
    if (desc) desc.innerText = "第一步：在 PPT 复制图片 (Ctrl+C)\n第二步：点一下这里，按 Ctrl+V";
    
    // 2. 监听全局粘贴事件 (这是核心！无需权限即可触发)
    document.addEventListener("paste", handlePaste);
});

async function handlePaste(event) {
    // 阻止默认粘贴行为（防止它试图把图贴到文字里）
    event.preventDefault();
    
    updateStatus("⚡ 检测到粘贴！正在处理...");

    // 1. 从粘贴事件中获取数据
    const items = (event.clipboardData || event.originalEvent.clipboardData).items;
    let blob = null;

    // 2. 寻找图片
    for (const item of items) {
        if (item.type.indexOf("image") === 0) {
            blob = item.getAsFile();
            break;
        }
    }

    if (!blob) {
        updateStatus("❌ 你粘贴的不是图片！\n请先在 PPT 里选中图片复制。");
        return;
    }

    try {
        // 3. 将 Blob 转为 Base64
        const base64 = await blobToBase64(blob);
        
        updateStatus("🎨 正在进行反色计算...");

        // 4. 反色处理
        const newBase64 = await invertImagePromise(base64);

        // 5. 将结果写回剪贴板
        // 注意：写入剪贴板通常比读取要宽松，但为了保险，我们需要一个 Blob
        const newBlob = await base64ToBlob(newBase64);
        
        await navigator.clipboard.write([
            new ClipboardItem({ [blob.type]: newBlob })
        ]);

        updateStatus("✅ 成功！新图已复制。\n请回到 PPT 按 Ctrl+V");
        
        // 视觉反馈：让按钮变绿一下
        const btn = document.getElementById("runBtn");
        if(btn) {
            const oldText = btn.innerText;
            btn.style.backgroundColor = "#107c10";
            btn.innerText = "完成！请粘贴";
            setTimeout(() => {
                btn.style.backgroundColor = ""; // 恢复颜色
                btn.innerText = oldText;
            }, 3000);
        }

    } catch (err) {
        console.error(err);
        updateStatus("⚠️ 处理出错: " + err.message);
    }
}

// --- 辅助函数：Blob 转 Base64 ---
function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

// --- 辅助函数：Base64 转 Blob ---
async function base64ToBlob(base64) {
    const res = await fetch(base64);
    return await res.blob();
}

// --- 核心算法：反色 ---
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
            // RGB 反色
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
