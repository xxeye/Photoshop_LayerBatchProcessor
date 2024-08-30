#target photoshop

function main() {
    var doc = app.activeDocument;

    // 讓用戶選擇存儲位置
    var outputFolder = Folder.selectDialog("選擇儲存 PNG 圖片的資料夾");

    if (outputFolder == null) {
        alert("未選擇儲存位置，腳本已取消。");
        return;
    }

    // 另存副本以避免改動原始文件
    var tempFile = new File(outputFolder.fsName + "/temp_copy.psd");
    var tempDoc = doc.duplicate();
    tempDoc.saveAs(tempFile);

    // 在步驟1前刪除名稱含"!"，和不可見的頂層圖層組，包含其下子集一併刪除
    deleteExclusionGroups(tempDoc);

    // 合併名稱中含有 "#" 的圖層組 並 同時將同層級的 "$" 圖層複製到 "@" 開頭的圖層組內，並刪除名稱中的 "$" 符號
    mergeHashAndProcessDollarLayers(tempDoc);

    // 遍歷嵌套結構並修改圖層組名稱，接著將包含 "@" 的組移至頂層並刪除名稱中的 "@"
    renameAndMoveAtGroups(tempDoc);

    // 移除空白圖層組
    removeEmptyGroups(tempDoc);

    // 獲取頂層圖層組並保存為 PNG
    var topLevelGroups = getTopLevelGroups(tempDoc);

    for (var i = 0; i < topLevelGroups.length; i++) {
        var group = topLevelGroups[i];
        
        // 將圖層組轉為智慧型物件
        var smartObject = convertToSmartObject(group);

        // 開啟智慧型物件
        openSmartObject(smartObject);

                // 检查内容是否为空
        if (!hasVisiblePixels(app.activeDocument)) {
            // 如果内容为空，关闭智慧型物件但不保存更改
            app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
            // 返回副本文件继续处理其他图层
            app.activeDocument = tempDoc;
            continue; // 跳过保存此图层组
        }

        // 調整圖像大小
        adjustCanvasSize(app.activeDocument);

        // 儲存為 PNG
        var fileName = smartObject.name + ".png";
        var folder = determineFolder(outputFolder, smartObject.name);

        saveAsImage(app.activeDocument, folder, fileName);

        // 關閉智慧型物件並儲存更改
        app.activeDocument.close(SaveOptions.SAVECHANGES);

        // 返回副本文件
        app.activeDocument = tempDoc;
    }

    // 關閉副本文件且不保存更改
    tempDoc.close(SaveOptions.DONOTSAVECHANGES);

    // 刪除副本文件
    try {
        tempFile.remove();
    } catch (e) {
        alert("無法刪除臨時文件：" + e.message);
    }

    // 彈出完成提示
    alert("已完成");
}

function hasVisiblePixels(doc) {
    var layers = doc.layers;
    for (var i = 0; i < layers.length; i++) {
        if (layers[i].visible && layers[i].bounds[2] > layers[i].bounds[0] && layers[i].bounds[3] > layers[i].bounds[1]) {
            return true; // 有可见像素
        }
    }
    return false; // 没有可见像素
}

// 刪除名稱含"!"和不可見的頂層圖層組，包含其下子集一併刪除
function deleteExclusionGroups(doc) {
    var groupsToDelete = [];

    // 首先遍历一遍，找出所有需要删除的图层组
    for (var i = 0; i < doc.layerSets.length; i++) {
        var group = doc.layerSets[i];
        if (group.name.indexOf("!") !== -1 || !group.visible) {
            groupsToDelete.push(group);
        }
    }

    // 然后再遍历一遍，将需要删除的图层组全部删除
    for (var i = 0; i < groupsToDelete.length; i++) {
        groupsToDelete[i].remove();
    }
}


// 合併名稱中含有 "#" 的圖層組 並 同時將同層級的 "$" 圖層複製到 "@" 開頭的圖層組內，並刪除名稱中的 "$" 符號
function mergeHashAndProcessDollarLayers(doc) {
    for (var i = 0; i < doc.layerSets.length; i++) {
        processLayerGroup(doc.layerSets[i]);
    }
}

function processLayerGroup(layerSet) {
    for (var i = 0; i < layerSet.layerSets.length; i++) {
        var group = layerSet.layerSets[i];

        // 合併名稱中含有 "#" 的圖層組
        if (group.name.indexOf("#") !== -1) {
            group.merge();
        } else {
            processLayerGroup(group);
        }

        // 處理 "$" 開頭的圖層
        var dollarLayer = null;
        for (var j = 0; j < layerSet.artLayers.length; j++) {
            var layer = layerSet.artLayers[j];
            if (layer.name.indexOf('$') === 0) {
                dollarLayer = layer;
                break;
            }
        }

        if (dollarLayer !== null) {
            var newLayerName = dollarLayer.name.replace('$', '');

            // 將 "$" 圖層複製到 "@" 開頭的圖層組內
            for (var k = 0; k < layerSet.layerSets.length; k++) {
                var subGroup = layerSet.layerSets[k];
                if (subGroup.name.indexOf('@') === 0) {
                    var copiedLayer = dollarLayer.duplicate();
                    copiedLayer.name = newLayerName;
                    copiedLayer.move(subGroup, ElementPlacement.INSIDE);
                }
            }

            // 刪除原始的 "$" 圖層
            dollarLayer.remove();
        }
    }
}

// 遍歷嵌套結構並修改圖層組名稱，接著將包含 "@" 的組移至頂層並刪除名稱中的 "@"
function renameAndMoveAtGroups(doc) {
    for (var i = 0; i < doc.layerSets.length; i++) {
        renameAndMoveNestedAtGroups(doc.layerSets[i], "");
    }
}

function renameAndMoveNestedAtGroups(layerSet, prefix) {
    for (var i = layerSet.layerSets.length - 1; i >= 0; i--) {
        var group = layerSet.layerSets[i];
        var newPrefix = prefix + layerSet.name + "_";

        if (group.name.indexOf("@") !== -1) {
            group.name = newPrefix + group.name.replace("@", "");
            group.move(app.activeDocument, ElementPlacement.PLACEATEND);
        }

        renameAndMoveNestedAtGroups(group, newPrefix);
    }
}

function removeEmptyGroups(layerSet) {
    // 遍歷圖層組內的子圖層組，進行遞歸刪除
    for (var i = layerSet.layerSets.length - 1; i >= 0; i--) {
        var group = layerSet.layerSets[i];
        removeEmptyGroups(group); // 遞歸呼叫自身以處理嵌套的圖層組
    }

    // 如果圖層組內沒有子圖層組且沒有像素圖層，則刪除該圖層組
    if (layerSet.layerSets.length === 0 && layerSet.artLayers.length === 0) {
        layerSet.remove();
    }
}

function getTopLevelGroups(doc) {
    var groups = [];
    for (var i = 0; i < doc.layerSets.length; i++) {
        groups.push(doc.layerSets[i]);
    }
    return groups;
}

function convertToSmartObject(layerSet) {
    app.activeDocument.activeLayer = layerSet;
    executeAction(stringIDToTypeID("newPlacedLayer"), undefined, DialogModes.NO);
    return app.activeDocument.activeLayer;
}

function openSmartObject(layer) {
    app.activeDocument.activeLayer = layer;
    executeAction(stringIDToTypeID("placedLayerEditContents"), undefined, DialogModes.NO);
}

function determineFolder(outputFolder, name) {
    var suffixes = ["CNY", "ENU", "VND", "THB", "ESP", "PTE", "CHT"];

    // 檢查是否為語言後綴並創建對應資料夾
    for (var i = 0; i < suffixes.length; i++) {
        if (name.substring(name.length - suffixes[i].length) === suffixes[i]) {
            var localeFolder = new Folder(outputFolder.fsName + "/Locale");
            if (!localeFolder.exists) localeFolder.create();
            
            var specificFolder = new Folder(localeFolder.fsName + "/" + suffixes[i]);
            if (!specificFolder.exists) specificFolder.create();
            return specificFolder;
        }
    }

    // 處理純數字名稱或名稱後綴為兩位或三位數字的檔案
    var nameWithoutSuffix = name.slice(0, -2);
    var suffix = name.slice(-2);

    // 如果名稱為純數字，直接放入 num_ 資料夾
    if (isNumeric(name) && (name.length === 2 || name.length === 3)) {
        var numFolder = new Folder(outputFolder.fsName + "/num_");
        if (!numFolder.exists) numFolder.create();
        return numFolder;
    }

    // 檢查名稱的後綴是否為兩位或三位數字
    if (isNumeric(suffix) && suffix.length === 2) {
        var numFolder = new Folder(outputFolder.fsName + "/num_");
        if (!numFolder.exists) numFolder.create();
        var subFolder = new Folder(numFolder.fsName + "/" + nameWithoutSuffix);
        if (!subFolder.exists) subFolder.create();
        return subFolder;
    }

    // 三位數字後綴處理
    nameWithoutSuffix = name.slice(0, -3);
    suffix = name.slice(-3);

    if (isNumeric(suffix) && suffix.length === 3) {
        var numFolder = new Folder(outputFolder.fsName + "/num_");
        if (!numFolder.exists) numFolder.create();
        var subFolder = new Folder(numFolder.fsName + "/" + nameWithoutSuffix);
        if (!subFolder.exists) subFolder.create();
        return subFolder;
    }

    // 如果不符合上述條件，則分類到 Common 資料夾
    var commonFolder = new Folder(outputFolder.fsName + "/Common");
    if (!commonFolder.exists) commonFolder.create();
    return commonFolder;
}

function isNumeric(name) {
    var num = parseInt(name, 10);
    return !isNaN(num) && (name.length === 2 || name.length === 3);
}

function adjustCanvasSize(doc) {
    var width = doc.width.value;
    var height = doc.height.value;

    if (width % 2 !== 0) {
        doc.resizeCanvas(UnitValue(width + 1, "px"), doc.height);
    }

    if (height % 2 !== 0) {
        doc.resizeCanvas(doc.width, UnitValue(height + 1, "px"));
    }
}

function saveAsImage(doc, folder, filename) {
    var file;

    if (filename.toLowerCase().indexOf(".jpg") !== -1) {
        // 移除 .jpg 後綴
        filename = filename.substring(0, filename.toLowerCase().lastIndexOf(".jpg"));
        file = new File(folder.fsName + "/" + filename + ".jpg");

        // 保存為 JPG 格式
        var jpgOptions = new JPEGSaveOptions();
        jpgOptions.quality = 12; // 設定 JPEG 品質為最高
        doc.saveAs(file, jpgOptions, true, Extension.LOWERCASE);
    } else {
        // 移除 .png 後綴
        if (filename.toLowerCase().indexOf(".png") !== -1) {
            filename = filename.substring(0, filename.toLowerCase().lastIndexOf(".png"));
        }
        file = new File(folder.fsName + "/" + filename + ".png");

        // 保存為 PNG 格式
        var pngOptions = new PNGSaveOptions();
        doc.saveAs(file, pngOptions, true, Extension.LOWERCASE);
    }
}

main();
