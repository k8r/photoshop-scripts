#target photoshop

/*
Exports PNGs from groups whose names start with "EX_".

For each visible EX_ group:
- Each visible immediate child subgroup prefixed "ex_" is exported as a PNG.
- If a sibling layer named "ws" exists, its bounds define the
  export canvas size for all PNGs in that group.
- If no "ws" layer exists, exports use the full document canvas size.
- Invisible EX_ groups and invisible ex_ subgroups are skipped.

Files are named after the subgroup and saved to a folder chosen when the
script runs.
*/

app.bringToFront();

(function () {
    if (!app.documents.length) {
        alert("Open a PSD first.");
        return;
    }

    var doc = app.activeDocument;

    var EXPORT_PREFIX = "EX_";
    var CHILD_EXPORT_PREFIX = "ex_";
    var WIDGET_SIZE_LAYER = "ws";
    var PATH_PREFIX = "path_";
    var HIDE_WIDGET_IN_EXPORT = true;

    var settingsFile = new File(File($.fileName).parent.fsName + "/export-ex-groups.lastdir");
    var startFolder = Folder.desktop;
    if (settingsFile.exists) {
        settingsFile.open("r");
        var savedPath = settingsFile.readln();
        settingsFile.close();
        var saved = new Folder(savedPath);
        if (saved.exists) startFolder = saved;
    }

    var outputFolder = startFolder.selectDlg("Choose an export folder");
    if (!outputFolder) return;

    settingsFile.open("w");
    settingsFile.writeln(outputFolder.fsName);
    settingsFile.close();

    var originalRulerUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;

    try {
        var exportedCount = 0;
        var notes = [];

        for (var i = 0; i < doc.layerSets.length; i++) {
            var parentGroup = doc.layerSets[i];

            if (!startsWith(parentGroup.name, EXPORT_PREFIX)) {
                continue;
            }

            if (!parentGroup.visible) {
                continue;
            }

            var widgetLayer = findImmediateArtLayerByName(parentGroup, WIDGET_SIZE_LAYER);

            var pathLayer = findImmediateArtLayerByPrefix(parentGroup, PATH_PREFIX);
            var groupOutputFolder = outputFolder;
            if (pathLayer) {
                var subdir = pathLayer.name.substring(PATH_PREFIX.length);
                groupOutputFolder = new Folder(outputFolder.fsName + "/" + subdir);
                if (!groupOutputFolder.exists) {
                    groupOutputFolder.create();
                }
            }

            var exportRect;
            var usingWidget = false;

            if (widgetLayer) {
                var widgetBounds = getBoundsPx(widgetLayer.bounds);
                exportRect = {
                    left: widgetBounds.left,
                    top: widgetBounds.top,
                    right: widgetBounds.right,
                    bottom: widgetBounds.bottom
                };
                usingWidget = true;
            } else {
                exportRect = {
                    left: 0,
                    top: 0,
                    right: doc.width.as("px"),
                    bottom: doc.height.as("px")
                };
                notes.push("Parent group '" + parentGroup.name + "' has no '" + WIDGET_SIZE_LAYER + "' layer; used full canvas size.");
            }

            var exportWidth = exportRect.right - exportRect.left;
            var exportHeight = exportRect.bottom - exportRect.top;

            if (exportWidth <= 0 || exportHeight <= 0) {
                notes.push("Skipped parent group '" + parentGroup.name + "' because export bounds were invalid.");
                continue;
            }

            for (var j = 0; j < parentGroup.layerSets.length; j++) {
                var childGroup = parentGroup.layerSets[j];

                if (!startsWith(childGroup.name, CHILD_EXPORT_PREFIX)) {
                    continue;
                }

                if (!childGroup.visible) {
                    continue;
                }

                exportChildGroup(
                    doc,
                    parentGroup,
                    childGroup,
                    widgetLayer,
                    exportWidth,
                    exportHeight,
                    usingWidget,
                    groupOutputFolder
                );

                exportedCount++;
            }
        }

        var summary = "Done.\nExported: " + exportedCount;
        if (notes.length) {
            summary += "\n\nNotes:\n- " + notes.join("\n- ");
        }
        alert(summary);

    } catch (e) {
        alert("Error: " + e.message + "\nLine: " + e.line);
    } finally {
        app.preferences.rulerUnits = originalRulerUnits;
    }

    function exportChildGroup(sourceDoc, parentGroup, childGroup, widgetLayer, exportWidth, exportHeight, usingWidget, outputFolder) {
        var tempDoc = app.documents.add(
            exportWidth,
            exportHeight,
            sourceDoc.resolution,
            childGroup.name,
            NewDocumentMode.RGB,
            DocumentFill.TRANSPARENT
        );

        app.activeDocument = sourceDoc;

        var dupChildGroup = childGroup.duplicate(tempDoc, ElementPlacement.PLACEATBEGINNING);
        var dupWidgetLayer = null;

        if (usingWidget && widgetLayer) {
            dupWidgetLayer = widgetLayer.duplicate(tempDoc, ElementPlacement.PLACEATBEGINNING);
        }

        app.activeDocument = tempDoc;

        if (usingWidget && dupWidgetLayer) {
            var dupWidgetBounds = getBoundsPx(dupWidgetLayer.bounds);
            var dx = -dupWidgetBounds.left;
            var dy = -dupWidgetBounds.top;

            dupChildGroup.translate(dx, dy);
            dupWidgetLayer.translate(dx, dy);

            if (HIDE_WIDGET_IN_EXPORT) {
                dupWidgetLayer.visible = false;
            }
        } else {
            tempDoc.trim(TrimType.TRANSPARENT);
        }

        var safeName = sanitizeFileName(childGroup.name.substring(CHILD_EXPORT_PREFIX.length));
        var outFile = new File(outputFolder.fsName + "/" + safeName + ".png");
        saveDocumentAsPNG(tempDoc, outFile);

        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = sourceDoc;
    }

    function findImmediateArtLayerByName(group, targetName) {
        for (var i = 0; i < group.artLayers.length; i++) {
            if (group.artLayers[i].name === targetName) {
                return group.artLayers[i];
            }
        }
        return null;
    }

    function findImmediateArtLayerByPrefix(group, prefix) {
        for (var i = 0; i < group.artLayers.length; i++) {
            if (startsWith(group.artLayers[i].name, prefix)) {
                return group.artLayers[i];
            }
        }
        return null;
    }

    function getBoundsPx(bounds) {
        return {
            left: bounds[0].as("px"),
            top: bounds[1].as("px"),
            right: bounds[2].as("px"),
            bottom: bounds[3].as("px")
        };
    }

    function saveDocumentAsPNG(documentRef, outFile) {
        var pngOptions = new PNGSaveOptions();
        pngOptions.compression = 9;
        pngOptions.interlaced = false;
        documentRef.saveAs(outFile, pngOptions, true, Extension.LOWERCASE);
    }

    function startsWith(str, prefix) {
        return str.indexOf(prefix) === 0;
    }

    function sanitizeFileName(name) {
        return name.replace(/[\\\/:*?"<>|]/g, "_");
    }
})();