// 双色替换-最终稳定版
#target Illustrator
app.executeMenuCommand('doc-color-cmyk');
var myDoc = app.activeDocument;

$.writeln("════════════════════════════");
$.writeln("开始执行时间：" + new Date());
$.writeln("════════════════════════════");

// 定义无颜色标记
var NO_COLOR = "NoColor";

// 定义颜色映射（显示用3位数，实际使用时解析）
var colorMap = {
    "080_074_072_048": "000_000_000_100",   // 标题-黑色文字
    "010_007_003_000": "000_000_000_015",   // 标题-灰色阴影
    "067_058_055_004": "000_000_000_070",   // 线条-灰色颜色
    "076_069_066_029": "000_000_000_100",   // 正文-黑色文字
    "072_016_001_000": "090_000_000_000",   // 主题-青色阴影
    
    "093_088_089_080": "000_000_000_100",   // 主题-其他重点
    "000_096_095_000": "100_000_000_000",   // 主题-其他重点
    "000_073_083_000": "100_000_000_000",   // 主题-其他重点
    "064_034_000_000": "100_000_000_000",   // 主题-其他重点
};

// 默认值配置
var defaultBlack = "000_000_000_100";  // 默认黑色
var defaultDual = "030_000_000_000";   // 默认双色（30%青）

// 定义跳过的颜色值（这些颜色不会被修改）
var skipColors = {
    "000_000_000_000": true,            // 白色
    "100_000_000_000": true,            // 100%青
    "000_100_000_000": true,            // 100%品红
    "000_000_100_000": true,            // 100%黄
    "000_000_000_100": true             // 100%黑
};

// 记录总数
var textCount = myDoc.textFrames.length;
var pathCount = myDoc.pathItems.length;

$.writeln("文字对象数：" + textCount);
$.writeln("图形对象数：" + pathCount);
$.writeln("────────────────────────────");

// 安全的获取颜色键值函数
function safeGetKey(colorObj) {
    if (!colorObj) return NO_COLOR;
    try {
        if (colorObj.typename == "CMYKColor") {
            var cVal = Math.round(colorObj.cyan);
            var mVal = Math.round(colorObj.magenta);
            var yVal = Math.round(colorObj.yellow);
            var kVal = Math.round(colorObj.black);
            
            return padZero(cVal, 3) + "_" + padZero(mVal, 3) + "_" + 
                   padZero(yVal, 3) + "_" + padZero(kVal, 3);
        }
    } catch (e) {
        $.writeln("获取颜色键值时出错：" + e.message);
    }
    return NO_COLOR;
}

// 获取颜色总数
function getColorCount(obj) {
    var count = 0;
    for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
            count++;
        }
    }
    return count;
}

// 补零函数
function padZero(num, length) {
    var str = num.toString();
    while (str.length < length) {
        str = "0" + str;
    }
    return str;
}

//PDF
function savePDF( file ) {
    var saveOpts = new PDFSaveOptions();
    saveOpts.compatibility = PDFCompatibility.ACROBAT4;//兼容PDF1.3
    saveOpts.preserveEditability = true;//保留AI编辑功能	
    saveOpts.generateThumbnails = true;//嵌入页面缩略图
    myDoc.saveAs( file, saveOpts );
}


function main(){
    // 统计文档中的颜色
    var colorStats = {};

    // 统计文字颜色
    for (var i = 0; i < textCount; i++) {
        var tf = myDoc.textFrames[i];
        var attrs = tf.textRange.characterAttributes;
        
        var fillKey = safeGetKey(attrs.fillColor);
        colorStats[fillKey] = (colorStats[fillKey] || 0) + 1;
        
        var strokeKey = safeGetKey(attrs.strokeColor);
        colorStats[strokeKey] = (colorStats[strokeKey] || 0) + 1;
    }

    // 统计图形颜色
    for (var i = 0; i < pathCount; i++) {
        var item = myDoc.pathItems[i];
        
        var fillKey = safeGetKey(item.fillColor);
        colorStats[fillKey] = (colorStats[fillKey] || 0) + 1;
        
        var strokeKey = safeGetKey(item.strokeColor);
        colorStats[strokeKey] = (colorStats[strokeKey] || 0) + 1;
    }

    // 检查匹配情况
    var missingMappings = [];
    var skipList = [];
    var matchList = [];
    var noColorCount = 0;

    for (var key in colorStats) {
        var target = colorMap[key];
        var count = colorStats[key];
        
        if (key == NO_COLOR) {
            noColorCount = count;
            continue;
        }
        
        if (skipColors[key]) {
            skipList.push({key: key, count: count});
            continue;
        }
        
        if (!target) {
            missingMappings.push({key: key, count: count});
            continue;
        }
        
        var display = "  ✓ " + key + " → " + target + " (使用" + count + "次)";
        matchList.push(display);
    }

    if (missingMappings.length > 0) {
        var errorMsg = "❌ 颜色匹配错误！\n";
        errorMsg += "════════════════════════════\n";
        errorMsg += "\n以下颜色在 colorMap 中缺失：\n";
        
        for (var i = 0; i < missingMappings.length; i++) {
            errorMsg += "  ✗ " + missingMappings[i].key + " (使用" + missingMappings[i].count + "次)\n";
        }
        
        errorMsg += "\n期望的颜色映射：\n";
        errorMsg += "  076_069_066_029 → 000_000_000_100\n";
        errorMsg += "  080_074_072_048 → 025_000_000_000\n";
        errorMsg += "  067_058_055_004 → 000_000_000_070\n";
        errorMsg += "  010_007_003_000 → 015_000_000_000\n";
        errorMsg += "  072_016_001_000 → 090_000_000_000\n";
        
        errorMsg += "\n请检查颜色值是否正确！";
        alert(errorMsg);
        return;
    }

    // 构建确认信息
    var confirmMsg = "📊 颜色匹配信息\n";
    confirmMsg += "════════════════════════════\n";
    confirmMsg += "总颜色：" + getColorCount(colorStats) + "个\n";
    confirmMsg += "已匹配：" + matchList.length + "个\n";
    confirmMsg += "已跳过：" + skipList.length + "个\n";
    confirmMsg += "无颜色：" + (noColorCount > 0 ? "是" : "否") + "\n\n";

    if (noColorCount > 0) {
        confirmMsg += "无颜色对象：\n";
        confirmMsg += "  ⏭️ " + NO_COLOR + " (不处理) (使用" + noColorCount + "次)\n";
        confirmMsg += "────────────────────────────\n";
    }

    if (skipList.length > 0) {
        confirmMsg += "跳过的颜色：\n";
        for (var i = 0; i < skipList.length; i++) {
            confirmMsg += "  ⏭️ " + skipList[i].key + " (跳过处理) (使用" + skipList[i].count + "次)\n";
        }
        confirmMsg += "────────────────────────────\n";
    }

    for (var i = 0; i < matchList.length; i++) {
        confirmMsg += matchList[i] + "\n";
    }

    confirmMsg += "\n默认值配置：\n";
    confirmMsg += "  • 默认黑色：" + defaultBlack + "\n";
    confirmMsg += "  • 默认双色：" + defaultDual + "\n";
    confirmMsg += "\n是否开始替换？";

    if (confirm(confirmMsg)) {
        try {
            // ===== 第一步：处理文字 =====
            $.writeln("\n开始处理文字对象...");
            
            for (var i = 0; i < textCount; i++) {
                var tf = myDoc.textFrames[i];
                var attrs = tf.textRange.characterAttributes;
                
                // 处理填充色
                try {
                    if (attrs.fillColor && attrs.fillColor.typename != "NoColor") {
                        var fillKey = safeGetKey(attrs.fillColor);
                        if (!skipColors[fillKey]) {
                            var target = colorMap[fillKey] || defaultDual;
                            var v = target.split("_");
                            var newColor = new CMYKColor();
                            newColor.cyan = parseInt(v[0], 10);
                            newColor.magenta = parseInt(v[1], 10);
                            newColor.yellow = parseInt(v[2], 10);
                            newColor.black = parseInt(v[3], 10);
                            attrs.fillColor = newColor;
                        }
                    }
                } catch (e) {
                    $.writeln("处理文字填充色时出错：" + e.message);
                }
                
                // 处理描边色
                try {
                    if (attrs.strokeColor && attrs.strokeColor.typename != "NoColor") {
                        var strokeKey = safeGetKey(attrs.strokeColor);
                        if (!skipColors[strokeKey]) {
                            var target2 = colorMap[strokeKey] || defaultDual;
                            var v2 = target2.split("_");
                            var newColor2 = new CMYKColor();
                            newColor2.cyan = parseInt(v2[0], 10);
                            newColor2.magenta = parseInt(v2[1], 10);
                            newColor2.yellow = parseInt(v2[2], 10);
                            newColor2.black = parseInt(v2[3], 10);
                            attrs.strokeColor = newColor2;
                        }
                    }
                } catch (e) {
                    $.writeln("处理文字描边色时出错：" + e.message);
                }
                
                if ((i+1) % 10 == 0) {
                    $.writeln("文字处理进度：" + (i+1) + "/" + textCount);
                }
            }
            $.writeln("✅ 文字处理完成");
            
            // ===== 第二步：处理图形 =====
            $.writeln("\n开始处理图形对象...");
            
            for (var i = 0; i < pathCount; i++) {
                var item = myDoc.pathItems[i];
                
                // 处理填充色
                try {
                    if (item.fillColor && item.fillColor.typename != "NoColor") {
                        var fillKey = safeGetKey(item.fillColor);
                        if (!skipColors[fillKey]) {
                            var target = colorMap[fillKey] || defaultDual;
                            var v = target.split("_");
                            var newColor = new CMYKColor();
                            newColor.cyan = parseInt(v[0], 10);
                            newColor.magenta = parseInt(v[1], 10);
                            newColor.yellow = parseInt(v[2], 10);
                            newColor.black = parseInt(v[3], 10);
                            item.fillColor = newColor;
                        }
                    }
                } catch (e) {
                    $.writeln("处理图形填充色时出错：" + e.message);
                }
                
                // 处理描边色
                try {
                    if (item.strokeColor && item.strokeColor.typename != "NoColor") {
                        var strokeKey = safeGetKey(item.strokeColor);
                        if (!skipColors[strokeKey]) {
                            var target2 = colorMap[strokeKey] || defaultDual;
                            var v2 = target2.split("_");
                            var newColor2 = new CMYKColor();
                            newColor2.cyan = parseInt(v2[0], 10);
                            newColor2.magenta = parseInt(v2[1], 10);
                            newColor2.yellow = parseInt(v2[2], 10);
                            newColor2.black = parseInt(v2[3], 10);
                            item.strokeColor = newColor2;
                        }
                    }
                } catch (e) {
                    $.writeln("处理图形描边色时出错：" + e.message);
                }
                
                if ((i+1) % 10 == 0) {
                    $.writeln("图形处理进度：" + (i+1) + "/" + pathCount);
                }
            }
            $.writeln("✅ 图形处理完成");
            
            // ===== 第三步：统计替换后的颜色 =====
            var afterStats = {};
            for (var i = 0; i < textCount; i++) {
                var tf = myDoc.textFrames[i];
                var attrs = tf.textRange.characterAttributes;
                
                var fillKey = safeGetKey(attrs.fillColor);
                afterStats[fillKey] = (afterStats[fillKey] || 0) + 1;
                
                var strokeKey = safeGetKey(attrs.strokeColor);
                afterStats[strokeKey] = (afterStats[strokeKey] || 0) + 1;
            }
            
            for (var i = 0; i < pathCount; i++) {
                var item = myDoc.pathItems[i];
                
                var fillKey = safeGetKey(item.fillColor);
                afterStats[fillKey] = (afterStats[fillKey] || 0) + 1;
                
                var strokeKey = safeGetKey(item.strokeColor);
                afterStats[strokeKey] = (afterStats[strokeKey] || 0) + 1;
            }
            
            // ===== 第四步：构建完成信息 =====
            var afterNoColor = 0;
            var afterAllColors = [];
            
            for (var key in afterStats) {
                if (key == NO_COLOR) {
                    afterNoColor = afterStats[key];
                } else {
                    afterAllColors.push({key: key, count: afterStats[key]});
                }
            }
            
            // 排序，让跳过的颜色在前面
            afterAllColors.sort(function(a, b) {
                var aSkip = skipColors[a.key] ? 0 : 1;
                var bSkip = skipColors[b.key] ? 0 : 1;
                return aSkip - bSkip;
            });
            
            // 构建弹窗信息
            var alertMsg = "✅ 替换完成！\n\n";
            alertMsg += "处理对象：" + (textCount + pathCount) + "个\n";
            alertMsg += "文字：" + textCount + "个\n";
            alertMsg += "图形：" + pathCount + "个\n\n";
            alertMsg += "📊 替换后的颜色统计：\n";
            alertMsg += "════════════════════════════\n";
            
            if (afterNoColor > 0) {
                alertMsg += "⏭️ " + NO_COLOR + "：" + afterNoColor + "次 (无颜色)\n";
            }
            
            for (var i = 0; i < afterAllColors.length; i++) {
                var prefix = skipColors[afterAllColors[i].key] ? "⏭️ " : "";
                alertMsg += prefix + afterAllColors[i].key + "：" + afterAllColors[i].count + "次\n";
            }

        //处理文字-转曲
        while (myDoc.textFrames.length != 0) {
            myDoc.textFrames[0].createOutline(); 
        }

        //另存为PDF
        var fpath = myDoc.path; 
        savePDF( fpath )

        alert(alertMsg);
            
        } catch (e) {
            $.writeln("\n❌ 出错位置：" + e.line);
            $.writeln("❌ 错误信息：" + e.message);
            alert("❌ 处理出错：\n" + e.message);
        }
    } else {
        alert("已取消替换");
    }
}

main();