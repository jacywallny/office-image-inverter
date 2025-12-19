Office.onReady((info) => {
    // 初始化界面逻辑
    const btn = document.getElementById("runBtn");
    if (btn) btn.onclick = runInvert;
});

async function runInvert() {
    updateStatus("⏳ 正在识别宿主环境...");

    // 👉 核心分流逻辑：你是 Word 还是 PPT？
    if (Office.context.host === Office.HostType.Word) {
        // 如果是 Word，走强力内核
        updateStatus("检测到 Word，启动强力读取模式...");
        await runInvertInWord();
    } else {
        // 如果是 PPT (或 Excel)，走通用兼容模式
        updateStatus("检测到 PowerPoint/Excel，启动通用模式...");
        runInvertCommon();
    }
}

// ==========================================
// 🔵 模式一：Word 专用强力内核 (你之前测试成功的那个)
// ==========================================
async function runInvertInWord() {
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

            updateStatus("🎨 Word: 读取成功，正在反色...");
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
// 🟠 模式二：PPT/通用兼容模式 (依靠旧版 API)
// ==========================================
function runInvertCommon() {
    // 尝试请求选区为“图片格式”
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Image, // 强行把选中的东西当图读
        { valueFormat: Office.ValueFormat.Base64 },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // PPT 这里最容易报错，所以要给出具体建议
                console.error(asyncResult.error);
                updateStatus("❌ PPT读取失败: " + asyncResult.error.message + 
                             "\n\n💡 提示：PPT 的 API 较弱，请确保：\n1. 只选中了一张图片\n2. 该图片不是组合形状");
            } else {
                const originalBase64 = asyncResult.value;
                updateStatus("🎨 PPT: 读取成功，正在反色...");
                
                invertImagePromise(originalBase64).then(newBase64 => {
                    const cleanBase64 = newBase64.split(",")[1];
                    
                    // 将新图片写回，替换当前选区
                    Office.context.document.setSelectedDataAsync(
                        cleanBase64,
                        { coercionType: Office.CoercionType.Image },
                        (res) => {
                            if (res.status === Office.AsyncResultStatus.Failed) {
                                updateStatus("❌ 替换失败: " + res.error.message);
                            } else {
                                updateStatus("✅ 成功！已反色");
                            }
                        }
                    );
                }).catch(err => {
                    updateStatus("⚠️ 处理错误: " + err);
                });
            }
        }
    );
}

// ==========================================
// 🎨 图像处理算法 (通用的)
// ==========================================
function invertImagePromise(base64Str) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        const prefix = "data:image/png;base64,";
        if (base64Str && !base64Str.startsWith("data:")) {
            img.src = prefix + base64Str;
        } else {
            img.src = base64Str;
        }

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
        img.onerror = (e) => reject(e);
    });
}

function updateStatus(message) {
    // 兼容之前的 UI 代码
    if (window.updateStatusUI) {
        window.updateStatusUI(message); // 如果你在 HTML 里写了 UI 逻辑
    } else {
        const el = document.getElementById("status");
        if(el) el.innerText = message;
    }
}
