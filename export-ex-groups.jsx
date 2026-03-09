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
        var skippedNoPrefix = [];
        var skippedHiddenParent = [];
        var skippedChildNoPrefix = [];
        var skippedHiddenChild = [];
        var groupBounds = [];
        var notes = [];

        for (var i = 0; i < doc.layerSets.length; i++) {
            var parentGroup = doc.layerSets[i];

            if (!startsWith(parentGroup.name, EXPORT_PREFIX)) {
                skippedNoPrefix.push(parentGroup.name);
                continue;
            }

            if (!parentGroup.visible) {
                skippedHiddenParent.push(parentGroup.name);
                continue;
            }

            var widgetSizeLayer = findImmediateArtLayerByName(parentGroup, WIDGET_SIZE_LAYER);

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

            if (widgetSizeLayer) {
                var wasVisible = widgetSizeLayer.visible;
                widgetSizeLayer.visible = true;
                var opaqueBounds = getCleanBounds(doc, widgetSizeLayer);
                widgetSizeLayer.visible = wasVisible;
                exportRect = {
                    left: opaqueBounds.left,
                    top: opaqueBounds.top,
                    right: opaqueBounds.right,
                    bottom: opaqueBounds.bottom
                };
                usingWidget = true;
            } else {
                exportRect = {
                    left: 0,
                    top: 0,
                    right: doc.width.as("px"),
                    bottom: doc.height.as("px")
                };
                notes.push("Parent group '" + parentGroup.name + "' has no '" + WIDGET_SIZE_LAYER + "' layer; will trim to visible content.");
            }

            var exportWidth = exportRect.right - exportRect.left;
            var exportHeight = exportRect.bottom - exportRect.top;

            groupBounds.push(parentGroup.name + ": " + exportWidth + "x" + exportHeight + " (" + exportRect.left + "," + exportRect.top + " -> " + exportRect.right + "," + exportRect.bottom + ")");

            if (exportWidth <= 0 || exportHeight <= 0) {
                notes.push("Skipped parent group '" + parentGroup.name + "' because export bounds were invalid.");
                continue;
            }

            for (var j = 0; j < parentGroup.layerSets.length; j++) {
                var childGroup = parentGroup.layerSets[j];

                if (!startsWith(childGroup.name, CHILD_EXPORT_PREFIX)) {
                    skippedChildNoPrefix.push(parentGroup.name + " > " + childGroup.name);
                    continue;
                }

                if (!childGroup.visible) {
                    skippedHiddenChild.push(parentGroup.name + " > " + childGroup.name);
                    continue;
                }

                exportChildGroup(
                    doc,
                    childGroup,
                    exportRect,
                    exportWidth,
                    exportHeight,
                    usingWidget,
                    groupOutputFolder
                );

                exportedCount++;
            }
        }

        var summary = "Done.\nExported: " + exportedCount;
        summary += "\nTop-level groups: " + doc.layerSets.length;
        if (groupBounds.length) {
            summary += "\n\nExport bounds:\n- " + groupBounds.join("\n- ");
        }
        if (skippedNoPrefix.length) {
            summary += "\n\nSkipped (no EX_ prefix):\n- " + skippedNoPrefix.join("\n- ");
        }
        if (skippedHiddenParent.length) {
            summary += "\n\nSkipped (hidden EX_ groups):\n- " + skippedHiddenParent.join("\n- ");
        }
        if (skippedChildNoPrefix.length) {
            summary += "\n\nSkipped (no ex_ prefix):\n- " + skippedChildNoPrefix.join("\n- ");
        }
        if (skippedHiddenChild.length) {
            summary += "\n\nSkipped (hidden ex_ subgroups):\n- " + skippedHiddenChild.join("\n- ");
        }
        if (notes.length) {
            summary += "\n\nNotes:\n- " + notes.join("\n- ");
        }
        alert(summary);

    } catch (e) {
        alert("Error: " + e.message + "\nLine: " + e.line);
    } finally {
        app.preferences.rulerUnits = originalRulerUnits;
    }

    function exportChildGroup(sourceDoc, childGroup, exportRect, exportWidth, exportHeight, usingWidget, outputFolder) {
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

        app.activeDocument = tempDoc;

        if (usingWidget) {
            dupChildGroup.translate(-exportRect.left, -exportRect.top);
        } else {
            tempDoc.trim(TrimType.TRANSPARENT);
        }

        var safeName = sanitizeFileName(childGroup.name.substring(CHILD_EXPORT_PREFIX.length));
        var outFile = new File(outputFolder.fsName + "/" + safeName + ".png");
        saveDocumentAsPNG(tempDoc, outFile);

        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = sourceDoc;
    }

    function cleanUpTransparency(docRef, layer) {
        // Delete all non-fully-opaque pixels from a layer.
        // This eliminates stray near-transparent pixels from Procreate exports.
        var prevActive = docRef.activeLayer;
        docRef.activeLayer = layer;

        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
        desc.putReference(charIDToTypeID("null"), ref);
        var ref2 = new ActionReference();
        ref2.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Trsp"));
        desc.putReference(charIDToTypeID("T   "), ref2);
        executeAction(charIDToTypeID("setd"), desc, DialogModes.NO);

        docRef.selection.invert();
        docRef.selection.clear();
        docRef.selection.deselect();

        docRef.activeLayer = prevActive;
    }

    function getCleanBounds(sourceDoc, layer) {
        // Clean up the layer, then get its bounds via a temp doc trim.
        var tempDoc = app.documents.add(
            sourceDoc.width,
            sourceDoc.height,
            sourceDoc.resolution,
            "bounds_check",
            NewDocumentMode.RGB,
            DocumentFill.TRANSPARENT
        );

        app.activeDocument = sourceDoc;
        var dupLayer = layer.duplicate(tempDoc, ElementPlacement.PLACEATBEGINNING);
        app.activeDocument = tempDoc;

        cleanUpTransparency(tempDoc, dupLayer);

        var fullWidth = tempDoc.width.as("px");
        var fullHeight = tempDoc.height.as("px");

        // Trim top and left to find content position
        tempDoc.trim(TrimType.TRANSPARENT, true, true, false, false);
        var left = fullWidth - tempDoc.width.as("px");
        var top = fullHeight - tempDoc.height.as("px");

        // Trim bottom and right to find content size
        tempDoc.trim(TrimType.TRANSPARENT, false, false, true, true);
        var contentWidth = tempDoc.width.as("px");
        var contentHeight = tempDoc.height.as("px");

        tempDoc.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = sourceDoc;

        return {
            left: left,
            top: top,
            right: left + contentWidth,
            bottom: top + contentHeight
        };
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