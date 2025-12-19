Office.onReady((info) => {
    const btn = document.getElementById("runBtn");
    if (btn) btn.onclick = runInvertByClipboard;
});

// ✂️ 剪贴板模式主函数
async function runInvertByClipboard() {
    updateStatus("⏳ 正在读取剪贴板...");

    try {
        // 1. 尝试从剪贴板读取内容
        // 注意：浏览器通常需要用户授权（第一次会弹窗）
        const clipboardItems = await navigator.clipboard.read();
        
        let foundImage = false;

        for (const item of clipboardItems) {
            // 2. 寻找图片格式 (png/jpeg)
            const imageType = item.types.find(type => type.startsWith("image/"));
            
            if (imageType) {
                foundImage = true;
                const blob = await item.getType(imageType);
                
                // 3. 将 Blob 转为 Base64 供我们处理
                const base64 = await blobToBase64(blob);
                
                updateStatus("🎨 获取成功，正在反色...");
                
                // 4. 反色处理
                const newBase64 = await invertImagePromise(base64);
                
                // 5. 将处理后的图片写回剪贴板
                const newBlob = await base64ToBlob(newBase64);
                
                // 写入剪贴板 (这就相当于你已经复制了新图)
                await navigator.clipboard.write([
                    new ClipboardItem({ [imageType]: newBlob })
                ]);
                
                updateStatus("✅ 成功！请按 Ctrl+V 粘贴");
                return; // 处理完一张就退出
            }
        }

        if (!foundImage) {
            updateStatus("❌ 剪贴板里没有图片！\n请先选中图片按 Ctrl+C");
        }

    } catch (err) {
        console.error(err);
        // 常见错误处理
        if (err.name === 'NotAllowedError') {
            updateStatus("❌ 权限被拒绝：请允许插件访问剪贴板");
        } else {
            updateStatus("⚠️ 错误: " + err.message + "\n请确保你先按了 Ctrl+C");
        }
    }
}

// --- 辅助工具：Blob 转 Base64 ---
function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

// --- 辅助工具：Base64 转 Blob ---
async function base64ToBlob(base64) {
    const res = await fetch(base64);
    return await res.blob();
}

// --- 图像处理核心算法 (不变) ---
function invertImagePromise(base64Str) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.src = base64Str;
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
        img.onerror = (e) => reject(e);
    });
}

function updateStatus(message) {
    const el = document.getElementById("status");
    if(el) el.innerText = message;
    
    // 如果你有美化版的 UI，这里适配一下颜色
    if (message.includes("Ctrl+V")) {
        if(el) el.style.color = "green";
        const btnText = document.getElementById("btnText");
        if(btnText) btnText.innerText = "已完成，请粘贴";
    }
}
