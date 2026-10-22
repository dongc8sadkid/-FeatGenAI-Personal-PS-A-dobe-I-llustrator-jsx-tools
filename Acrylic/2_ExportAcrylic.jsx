(function () {
    var PT_PER_INCH = 72;
    var PRE_CUT_MARGIN = 0.125 * PT_PER_INCH;
    var ORIENTATION_DOT_DIAMETER = 4.5;
    var LASER_STROKE_WIDTH = 1;

    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var setupDoc = app.activeDocument;

    if (setupDoc.artboards.length !== 1) {
        alert("Export stopped: the prepared File Setup document must contain exactly one artboard.");
        return;
    }

    var setupLaserLayer = findTopLevelLayer(setupDoc, "Lasercut");
    if (!setupLaserLayer || setupLaserLayer.pageItems.length === 0) {
        alert("Export stopped: no usable top-level Lasercut layer was found.");
        return;
    }

    var sourceArtboardRect = copyBounds(setupDoc.artboards[0].artboardRect);
    var sourceArtboardWidth = sourceArtboardRect[2] - sourceArtboardRect[0];
    var sourceArtboardHeight = sourceArtboardRect[1] - sourceArtboardRect[3];

    // Use the native Windows Save dialog instead of Folder.selectDialog so the
    // user can reach Quick Access and mapped drives directly. The chosen file
    // determines the output folder; the production filename remains fixed.
    var suggestedPDF = new File(
        getDefaultOutputFolder(setupDoc).fsName + "/File Setup.pdf"
    );
    var selectedPDF = suggestedPDF.saveDlg(
        "Save File Setup PDF (the other two files will use the same folder)",
        "Adobe PDF:*.pdf"
    );
    if (!selectedPDF) {
        return;
    }

    var outputFolder = selectedPDF.parent;
    var pdfFile = new File(outputFolder.fsName + "/File Setup.pdf");
    var laserFile = new File(outputFolder.fsName + "/+laser.ai");
    var preCutFile = new File(outputFolder.fsName + "/+Pre cut.ai");
    var outputFiles = [laserFile, preCutFile];

    // The native Save dialog already confirms replacement when its prefilled
    // File Setup.pdf target is used. If the user edits that filename, the script
    // still enforces File Setup.pdf and includes it in the custom overwrite check.
    if (!sameFilePath(selectedPDF, pdfFile)) {
        outputFiles.unshift(pdfFile);
    }
    var existingNames = getExistingFileNames(outputFiles);

    if (existingNames.length > 0) {
        var overwrite = confirm(
            "The following output file(s) already exist:\n\n" +
            existingNames.join("\n") +
            "\n\nOverwrite them?"
        );
        if (!overwrite) {
            return;
        }
    }

    // Remove replaceable outputs before any document save. If a target is the
    // currently open document, Illustrator will overwrite it through saveAs().
    try {
        removeExistingOutput(pdfFile, setupDoc);
        removeExistingOutput(laserFile, setupDoc);
        removeExistingOutput(preCutFile, setupDoc);
    } catch (prepareError) {
        alert("Export stopped:\n" + prepareError.message);
        return;
    }

    var tempPdfFile = new File(
        Folder.temp.fsName +
        "/Acrylic_File_Setup_" + (new Date().getTime()) + ".pdf"
    );
    var laserDoc = null;
    var preCutDoc = null;

    try {
        // 1) Full editable setup PDF, matching the default options shown in the
        // reference: Acrobat 7/PDF 1.6, editable AI data, thumbnails, and layers.
        setupDoc.saveAs(pdfFile, makePDFOptions());

        // 2) Work from a byte-for-byte temporary copy of the editable PDF. This
        // keeps the complete File Setup document open and untouched while the
        // laser-only CS6 variant is stripped down.
        if (tempPdfFile.exists) {
            tempPdfFile.remove();
        }
        if (!pdfFile.copy(tempPdfFile.fsName)) {
            throw new Error("Could not create the temporary editable PDF used for +laser.ai.");
        }

        laserDoc = app.open(tempPdfFile);
        unlockAndShowAllLayers(laserDoc);

        var laserOnlyLayer = findTopLevelLayer(laserDoc, "Lasercut");
        if (!laserOnlyLayer) {
            throw new Error("The temporary File Setup copy no longer contains a Lasercut layer.");
        }

        deleteAllTopLevelLayersExcept(laserDoc, laserOnlyLayer);
        laserOnlyLayer.locked = false;
        laserOnlyLayer.visible = true;
        laserDoc.saveAs(laserFile, makeCS6Options());
        laserDoc.close(SaveOptions.DONOTSAVECHANGES);
        laserDoc = null;

        if (tempPdfFile.exists && !tempPdfFile.remove()) {
            throw new Error("The temporary PDF could not be removed after +laser.ai was created.");
        }

        // 3) Build +Pre cut.ai as a fresh, single-layer CMYK CS6 document.
        preCutDoc = createPreCutDocument(sourceArtboardWidth, sourceArtboardHeight);
        preCutDoc.saveAs(preCutFile, makeCS6Options());
        preCutDoc.close(SaveOptions.DONOTSAVECHANGES);
        preCutDoc = null;

        setupDoc.activate();
        alert(
            "Acrylic export complete:\n\n" +
            "File Setup.pdf\n" +
            "+laser.ai\n" +
            "+Pre cut.ai"
        );
    } catch (error) {
        closeWithoutSaving(preCutDoc);
        closeWithoutSaving(laserDoc);

        if (tempPdfFile.exists) {
            try {
                tempPdfFile.remove();
            } catch (tempCleanupError) {
                // Keep the original export error as the useful message.
            }
        }

        try {
            setupDoc.activate();
        } catch (activateError) {
            // The output error below is still the useful result.
        }

        alert("Acrylic export stopped:\n" + error.message);
    }

    function makePDFOptions() {
        var options = new PDFSaveOptions();
        options.compatibility = PDFCompatibility.ACROBAT7;
        options.preserveEditability = true;
        options.generateThumbnails = true;
        options.acrobatLayers = true;
        options.optimization = false;
        options.viewAfterSaving = false;
        return options;
    }

    function makeCS6Options() {
        var options = new IllustratorSaveOptions();
        options.compatibility = Compatibility.ILLUSTRATOR16;
        options.compressed = true;
        options.pdfCompatible = true;
        return options;
    }

    function createPreCutDocument(rectangleWidth, rectangleHeight) {
        var artboardWidth = rectangleWidth + 2 * PRE_CUT_MARGIN;
        var artboardHeight = rectangleHeight + 2 * PRE_CUT_MARGIN;
        var preset = new DocumentPreset();
        preset.title = "+Pre cut";
        preset.width = artboardWidth;
        preset.height = artboardHeight;
        preset.units = RulerUnits.Inches;
        preset.colorMode = DocumentColorSpace.CMYK;
        preset.numArtboards = 1;

        // Creating the document through an inch-based preset prevents this
        // export from changing Illustrator's next New Document unit to Points.
        var document = app.documents.addDocument(
            getStartupPresetName(),
            preset,
            false
        );

        var layer = document.layers[0];
        layer.name = "Lasercut";
        layer.locked = false;
        layer.visible = true;

        var artboardRect = document.artboards[0].artboardRect;
        var artboardCenterX = (artboardRect[0] + artboardRect[2]) / 2;
        var artboardCenterY = (artboardRect[1] + artboardRect[3]) / 2;
        var rectangleLeft = artboardCenterX - rectangleWidth / 2;
        var rectangleTop = artboardCenterY + rectangleHeight / 2;
        var laserColor = makeCMYK(99, 0, 0, 0);

        var rectangle = layer.pathItems.rectangle(
            rectangleTop,
            rectangleLeft,
            rectangleWidth,
            rectangleHeight
        );
        styleLaserPath(rectangle, laserColor);
        rectangle.name = "Pre Cut Rectangle";

        var orientationDot = layer.pathItems.ellipse(
            artboardRect[1],
            artboardRect[2] - ORIENTATION_DOT_DIAMETER,
            ORIENTATION_DOT_DIAMETER,
            ORIENTATION_DOT_DIAMETER
        );
        styleLaserPath(orientationDot, laserColor);
        orientationDot.name = "Pre Cut Orientation Dot";

        return document;
    }

    function getStartupPresetName() {
        var presets = app.startupPresetsList;
        if (!presets || presets.length === 0) {
            throw new Error("Illustrator has no startup preset available for +Pre cut.ai.");
        }

        // Every size, color mode, artboard count, and ruler unit used by this
        // file is overridden above, so the first installed preset is safe and
        // also works with localized Illustrator installations.
        return presets[0];
    }

    function styleLaserPath(pathItem, laserColor) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeColor = laserColor;
        pathItem.strokeWidth = LASER_STROKE_WIDTH;
        pathItem.strokeOverprint = false;
    }

    function findTopLevelLayer(document, name) {
        var i;
        for (i = 0; i < document.layers.length; i++) {
            if (document.layers[i].name === name) {
                return document.layers[i];
            }
        }
        return null;
    }

    function unlockAndShowAllLayers(document) {
        var i;
        for (i = 0; i < document.layers.length; i++) {
            unlockAndShowLayerTree(document.layers[i]);
        }
    }

    function unlockAndShowLayerTree(layer) {
        layer.locked = false;
        layer.visible = true;

        var i;
        for (i = 0; i < layer.layers.length; i++) {
            unlockAndShowLayerTree(layer.layers[i]);
        }
    }

    function deleteAllTopLevelLayersExcept(document, layerToKeep) {
        var i;
        for (i = document.layers.length - 1; i >= 0; i--) {
            if (document.layers[i] !== layerToKeep) {
                document.layers[i].locked = false;
                document.layers[i].visible = true;
                document.layers[i].remove();
            }
        }
    }

    function getExistingFileNames(files) {
        var names = [];
        var i;
        for (i = 0; i < files.length; i++) {
            if (files[i].exists) {
                names.push(files[i].name);
            }
        }
        return names;
    }

    function getDefaultOutputFolder(document) {
        try {
            return document.fullName.parent;
        } catch (error) {
            return Folder.desktop;
        }
    }

    function removeExistingOutput(file, currentDocument) {
        if (!file.exists || isDocumentFile(currentDocument, file)) {
            return;
        }

        if (!file.remove()) {
            throw new Error("Could not replace the existing file: " + file.name);
        }
    }

    function isDocumentFile(document, file) {
        try {
            return sameFilePath(document.fullName, file);
        } catch (error) {
            return false;
        }
    }

    function sameFilePath(firstFile, secondFile) {
        return normalizePath(firstFile.fsName) === normalizePath(secondFile.fsName);
    }

    function normalizePath(path) {
        return path.toLowerCase().replace(/\\/g, "/");
    }

    function closeWithoutSaving(document) {
        if (!document) {
            return;
        }

        try {
            document.close(SaveOptions.DONOTSAVECHANGES);
        } catch (error) {
            // Ignore cleanup-only failures.
        }
    }

    function makeCMYK(cyan, magenta, yellow, black) {
        var color = new CMYKColor();
        color.cyan = cyan;
        color.magenta = magenta;
        color.yellow = yellow;
        color.black = black;
        return color;
    }

    function copyBounds(bounds) {
        return [bounds[0], bounds[1], bounds[2], bounds[3]];
    }
}());
