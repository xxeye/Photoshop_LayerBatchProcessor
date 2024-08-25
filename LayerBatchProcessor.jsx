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

    // 遍歷所有圖層組，將名稱含有 "#" 的圖層組合併圖層
    mergeHashGroups(tempDoc);

    // 遍歷嵌套結構並修改圖層組名稱
    renameAtGroups(tempDoc);

    // 遍歷所有圖層組，將包含 "@" 的組移至頂層並刪除名稱中的 "@"
    moveAndRenameAtGroups(tempDoc);

    // 移除空白圖層組
    removeEmptyGroups(tempDoc);

    // 獲取頂層圖層組
    var topLevelGroups = getTopLevelGroups(tempDoc);

    for (var i = 0; i < topLevelGroups.length; i++) {
        var group = topLevelGroups[i];
        
        // 跳過名稱中包含 "!" 的資料夾
        if (group.name.indexOf("!") !== -1) {
            continue;
        }
        
        // 將圖層組轉為智慧型物件
        var smartObject = convertToSmartObject(group);

        // 開啟智慧型物件
        openSmartObject(smartObject);

        // 調整圖像大小
        adjustCanvasSize(app.activeDocument);

        // 儲存為 PNG
        var fileName = smartObject.name + ".png";
        var folder = determineFolder(outputFolder, smartObject.name);

        saveAsPNG(app.activeDocument, folder, fileName);

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

function mergeHashGroups(doc) {
    for (var i = 0; i < doc.layerSets.length; i++) {
        mergeNestedHashGroups(doc.layerSets[i]);
    }
}

function mergeNestedHashGroups(layerSet) {
    for (var i = 0; i < layerSet.layerSets.length; i++) {
        var group = layerSet.layerSets[i];
        if (group.name.indexOf("#") !== -1) {
            group.merge();
        } else {
            mergeNestedHashGroups(group);
        }
    }
}

function renameAtGroups(doc) {
    for (var i = 0; i < doc.layerSets.length; i++) {
        renameNestedAtGroups(doc.layerSets[i], "");
    }
}

function renameNestedAtGroups(layerSet, prefix) {
    for (var i = 0; i < layerSet.layerSets.length; i++) {
        var group = layerSet.layerSets[i];
        var newPrefix = prefix + layerSet.name + "_";

        if (group.name.indexOf("@") !== -1) {
            group.name = newPrefix + group.name;
        }

        // 递归处理嵌套的组
        renameNestedAtGroups(group, newPrefix);
    }
}

function moveAndRenameAtGroups(doc) {
    for (var i = 0; i < doc.layerSets.length; i++) {
        moveAndRenameNestedAtGroups(doc.layerSets[i]);
    }
}

function moveAndRenameNestedAtGroups(layerSet) {
    for (var i = layerSet.layerSets.length - 1; i >= 0; i--) {
        var group = layerSet.layerSets[i];
        if (group.name.indexOf("@") !== -1) {
            group.move(layerSet, ElementPlacement.PLACEBEFORE);
            group.move(app.activeDocument, ElementPlacement.PLACEATEND);
            group.name = group.name.replace("@", ""); // 刪除名稱中的 "@"
        } else {
            moveAndRenameNestedAtGroups(group);
        }
    }
}

function removeEmptyGroups(doc) {
    for (var i = doc.layerSets.length - 1; i >= 0; i--) {
        var group = doc.layerSets[i];
        if (group.layerSets.length === 0 && group.artLayers.length === 0) {
            group.remove();
        }
    }
}

function getTopLevelGroups(doc) {
    var groups = [];
    for (var i = 0; i < doc.layerSets.length; i++) {
        var group = doc.layerSets[i];
        // 跳過名稱中包含 "!" 和不可見的資料夾
        if (group.name.indexOf("!") === -1 && group.visible) {
            groups.push(group);
        }
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
    var localeFolder = new Folder(outputFolder.fsName + "/Locale");
    if (!localeFolder.exists) localeFolder.create();

    for (var i = 0; i < suffixes.length; i++) {
        if (name.substring(name.length - suffixes[i].length) === suffixes[i]) {
            var specificFolder = new Folder(localeFolder.fsName + "/" + suffixes[i]);
            if (!specificFolder.exists) specificFolder.create();
            return specificFolder;
        }
    }

    var numFolder = new Folder(outputFolder.fsName + "/num_");
    if (isNumeric(name)) {
        if (!numFolder.exists) numFolder.create();
        return numFolder;
    }

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

function saveAsPNG(doc, folder, filename) {
    var file = new File(folder.fsName + "/" + filename);
    var pngOptions = new PNGSaveOptions();
    doc.saveAs(file, pngOptions, true, Extension.LOWERCASE);
}

main();
