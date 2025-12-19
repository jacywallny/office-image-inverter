Office.onReady((info) => {
    const btn = document.getElementById("runBtn");
    if (btn) btn.onclick = runInvert;
});

async function runInvert() {
    updateStatus("⏳ 正在处理...");
    
    // 判断环境
    if (Office.context.host === Office.HostType.Word) {
        await runInvertInWord();
    } else {
        runInvertCommon();
    }
}

// --- Word 专用强力模式 (修复版) ---
async function runInvertInWord() {
    try {
        await Word.run(async (context) => {
            // 1. 获取选区
            const selection = context.document.getSelection();
            const pictures = selection.inlinePictures;
            
            // 2. 加载图片列表
            pictures.load("items");
            await context.sync();

            if (pictures.items.length === 0) {
                updateStatus("❌ 未检测到图片！\n请右键图片 -> 自动换行 -> 设为【嵌入型】");
                return;
            }

            // 3. 拿到第一张图对象
            const wordPicture = pictures.items[0];

            // 【关键修改】使用方法来获取 Base64，而不是属性
            const base64Result = wordPicture.getBase64ImageSrc();
            
            // 必须再次同步，才能拿到方法返回的结果
            await context.sync();

            // 4. 提取数据
            const base64 = base64Result.value;
            if (!base64) {
                updateStatus("❌ 无法读取图片数据");
                return;
            }

            updateStatus("🎨 读取成功，正在反色...");

            // 5. 进行反色计算
            const newBase64 = await invertImagePromise(base64);

            // 6. 替换图片
            // 去掉前缀，只要数据部分
            const cleanBase64 = newBase64.split(",")[1];
            wordPicture.insertInlinePictureFromBase64(cleanBase64, "Replace");

            await context.sync();
            updateStatus("✅ 成功！已反色");
        });
    } catch (error) {
        console.error(error);
        updateStatus("⚠️ 报错: " + error.message);
    }
}

// --- PPT/通用模式 ---
function runInvertCommon() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Image,
        { valueFormat: Office.ValueFormat.Base64 },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                updateStatus("❌ 通用读取失败: " + asyncResult.error.message);
            } else {
                invertImagePromise(asyncResult.value).then(newBase64 => {
                    const cleanBase64 = newBase64.split(",")[1];
                    Office.context.document.setSelectedDataAsync(
                        cleanBase64,
                        { coercionType: Office.CoercionType.Image },
                        (res) => {
                            if (res.status === Office.AsyncResultStatus.Failed) updateStatus("替换失败");
                            else updateStatus("成功！");
                        }
                    );
                });
            }
        }
    );
}

// --- 图像处理核心算法 ---
function invertImagePromise(base64Str) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        // 兼容处理：有些返回带前缀，有些不带
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
            // 像素反色
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
}