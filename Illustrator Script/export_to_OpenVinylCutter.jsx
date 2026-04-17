/*
OpenVinylCutter Illustrator Bridge
Exports the current Illustrator selection to SVG and opens it in OpenVinylCutter.
*/

function ensureFolder(folder) {
    if (!folder.exists) {
        folder.create();
    }
}

function escapeForCmd(value) {
    return value.replace(/"/g, '""');
}

function getConfigFile() {
    var scriptFolder = new File($.fileName).parent;
    return new File(scriptFolder.fsName + "/OpenVinylCutter_config.txt");
}

function readSavedFolder() {
    var configFile = getConfigFile();
    if (!configFile.exists) {
        return null;
    }

    try {
        configFile.encoding = "UTF-8";
        configFile.open("r");
        var savedPath = configFile.readln();
        configFile.close();
        if (!savedPath) {
            return null;
        }
        return new Folder(savedPath);
    } catch (e) {
        try {
            configFile.close();
        } catch (_ignored) {
        }
        return null;
    }
}

function saveFolderPath(folder) {
    var configFile = getConfigFile();
    configFile.encoding = "UTF-8";
    configFile.lineFeed = "Windows";
    configFile.open("w");
    configFile.write(folder.fsName);
    configFile.close();
}

function hasOpenVinylCutterFiles(folder) {
    if (!folder || !folder.exists) {
        return false;
    }
    var launcher = new File(folder.fsName + "/open_svg.bat");
    return launcher.exists;
}

function findOpenVinylCutterFolder() {
    var scriptFile = new File($.fileName);
    var savedFolder = readSavedFolder();
    var candidates = [
        savedFolder,
        scriptFile.parent,
        Folder.desktop,
        new Folder(Folder.desktop.fsName + "/OpenVinylCutter"),
        new Folder(Folder.desktop.fsName + "/MAPJEE")
    ];

    for (var i = 0; i < candidates.length; i++) {
        if (hasOpenVinylCutterFiles(candidates[i])) {
            saveFolderPath(candidates[i]);
            return candidates[i];
        }
    }

    var selected = Folder.selectDialog("Select your OpenVinylCutter folder");
    if (hasOpenVinylCutterFiles(selected)) {
        saveFolderPath(selected);
        return selected;
    }

    return null;
}

function timestampString() {
    var now = new Date();
    function pad(value) {
        return (value < 10 ? "0" : "") + value;
    }
    return now.getFullYear() +
        pad(now.getMonth() + 1) +
        pad(now.getDate()) + "_" +
        pad(now.getHours()) +
        pad(now.getMinutes()) +
        pad(now.getSeconds());
}

function main() {
    if (app.documents.length === 0) {
        alert("Open a document first.");
        return;
    }

    var sourceDoc = app.activeDocument;
    if (!sourceDoc.selection || sourceDoc.selection.length === 0) {
        alert("Select one or more shapes first.");
        return;
    }

    var projectFolder = findOpenVinylCutterFolder();
    if (projectFolder === null) {
        alert("Could not find the OpenVinylCutter folder.\n\nMake sure open_svg.bat exists in your OpenVinylCutter project folder.");
        return;
    }
    var exportFolder = new Folder(projectFolder.fsName + "/bridge_exports");
    ensureFolder(exportFolder);

    var exportFile = new File(exportFolder.fsName + "/ovc_export_" + timestampString() + ".svg");
    var launcherFile = new File(projectFolder.fsName + "/open_svg.bat");

    var tempDoc = app.documents.add(DocumentColorSpace.RGB, sourceDoc.width, sourceDoc.height);
    tempDoc.activate();

    var copiedItems = [];
    for (var i = 0; i < sourceDoc.selection.length; i++) {
        copiedItems.push(sourceDoc.selection[i].duplicate(tempDoc, ElementPlacement.PLACEATEND));
    }

    tempDoc.selection = null;
    for (var j = 0; j < copiedItems.length; j++) {
        copiedItems[j].selected = true;
    }

    try {
        tempDoc.fitArtboardToSelectedArt(0);
    } catch (e) {
    }

    var options = new ExportOptionsSVG();
    options.embedRasterImages = true;
    options.fontSubsetting = SVGFontSubsetting.None;
    options.documentEncoding = SVGDocumentEncoding.UTF8;
    options.cssProperties = SVGCSSPropertyLocation.STYLEELEMENTS;
    options.coordinatePrecision = 3;
    options.fontType = SVGFontType.OUTLINEFONT;
    options.preserveEditability = false;

    tempDoc.exportFile(exportFile, ExportType.SVG, options);
    tempDoc.close(SaveOptions.DONOTSAVECHANGES);

    if (!launcherFile.exists) {
        alert("OpenVinylCutter launcher not found:\n" + launcherFile.fsName + "\n\nSVG was exported to:\n" + exportFile.fsName);
        return;
    }

    var tempLaunchFile = new File(exportFolder.fsName + "/launch_openvinylcutter.bat");
    tempLaunchFile.encoding = "UTF-8";
    tempLaunchFile.lineFeed = "Windows";
    tempLaunchFile.open("w");
    tempLaunchFile.write('@echo off\r\n');
    tempLaunchFile.write('call "' + escapeForCmd(launcherFile.fsName) + '" "' + escapeForCmd(exportFile.fsName) + '"\r\n');
    tempLaunchFile.close();
    tempLaunchFile.execute();
}

main();
