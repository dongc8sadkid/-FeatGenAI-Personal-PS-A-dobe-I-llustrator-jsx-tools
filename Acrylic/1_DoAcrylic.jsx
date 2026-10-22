(function () {
    var PT_PER_INCH = 72;
    var MAX_REGISTER_WIDTH = 47.75 * PT_PER_INCH;
    var ARTBOARD_MARGIN = 0.125 * PT_PER_INCH;
    var PRODUCT_TO_REGISTER_GAP = 0.25 * PT_PER_INCH;
    var TARGET_ITEM_GAP = 0.25 * PT_PER_INCH;

    var REGISTER_SHORT = 0.75 * PT_PER_INCH;
    var REGISTER_LONG = 1.75 * PT_PER_INCH;
    var CUT_SHORT = 0.5 * PT_PER_INCH;
    var CUT_LONG = 1.5 * PT_PER_INCH;
    var BLEED = 0.125 * PT_PER_INCH;

    var SMALL_CUT_STROKE = 1;
    var ORIENTATION_DOT_DIAMETER = 4.5;
    var EPSILON = 0.01;

    if (app.documents.length === 0) {
        alert("No document is open.");
        return;
    }

    var doc = app.activeDocument;
    var currentSelection = doc.selection;

    if (!currentSelection || currentSelection.length === 0) {
        alert("Select one finished laser-cut path, or a nested group of paths, then run the script again.");
        return;
    }

    // Keep stable references to the original selection. Creating new artwork
    // must not change which finished paths are restyled and moved at the end.
    var selection = [];
    var selectionIndex;
    for (selectionIndex = 0; selectionIndex < currentSelection.length; selectionIndex++) {
        selection.push(currentSelection[selectionIndex]);
    }

    var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
    var oldArtboardRect = copyBounds(doc.artboards[activeArtboardIndex].artboardRect);
    var createdItems = [];
    var createdLayers = [];
    var changedArtboard = false;
    var layerStates = [];

    try {
        var productBounds = getSelectionBounds(selection);
        var productWidth = productBounds[2] - productBounds[0];
        var productHeight = productBounds[1] - productBounds[3];

        validateSelectedLaserArtwork(selection);

        if (productWidth <= EPSILON || productHeight <= EPSILON) {
            throw new Error("The selected artwork must have both width and height.");
        }

        if (productWidth > MAX_REGISTER_WIDTH + EPSILON) {
            throw new Error(
                "The selected finished artwork is " + formatInches(productWidth) +
                " in wide. It must not exceed 47.75 in."
            );
        }

        var leftRightRegisterWidth = productWidth +
            2 * (PRODUCT_TO_REGISTER_GAP + REGISTER_SHORT);
        var useLeftRight = leftRightRegisterWidth <= MAX_REGISTER_WIDTH + EPSILON;

        // Create Register first and Lasercut second when both are missing so the
        // cut layer is naturally above the fill layer in a new document.
        var registerLayer = getOrCreateLayer(doc, "Register", createdLayers);
        var laserLayer = getOrCreateLayer(doc, "Lasercut", createdLayers);
        prepareLayer(registerLayer, layerStates);
        prepareLayer(laserLayer, layerStates);

        var registerColor = makeCMYK(0, 99, 0, 0);
        var laserColor = makeCMYK(99, 0, 0, 0);
        var allPairs = [];
        var orientationGapCreated = false;
        var shortEdgeWarning = false;

        if (useLeftRight) {
            var verticalDistribution = calculateDistribution(
                productHeight,
                REGISTER_LONG,
                TARGET_ITEM_GAP
            );

            var leftRegisterLeft = productBounds[0] -
                PRODUCT_TO_REGISTER_GAP - REGISTER_SHORT;
            var rightRegisterLeft = productBounds[2] + PRODUCT_TO_REGISTER_GAP;
            var leftPairs = [];
            var rightPairs = [];
            var i;

            for (i = 0; i < verticalDistribution.count; i++) {
                var registerTop = distributionPosition(
                    productBounds[1],
                    productHeight,
                    REGISTER_LONG,
                    verticalDistribution,
                    i,
                    true
                );

                var leftPair = createPair(
                    registerLayer,
                    laserLayer,
                    registerTop,
                    leftRegisterLeft,
                    REGISTER_SHORT,
                    REGISTER_LONG,
                    CUT_SHORT,
                    CUT_LONG,
                    registerColor,
                    laserColor,
                    "Left " + (i + 1),
                    createdItems
                );
                var rightPair = createPair(
                    registerLayer,
                    laserLayer,
                    registerTop,
                    rightRegisterLeft,
                    REGISTER_SHORT,
                    REGISTER_LONG,
                    CUT_SHORT,
                    CUT_LONG,
                    registerColor,
                    laserColor,
                    "Right " + (i + 1),
                    createdItems
                );

                leftPairs.push(leftPair);
                rightPairs.push(rightPair);
                allPairs.push(leftPair);
                allPairs.push(rightPair);
            }

            // Direction key: on the dot-adjacent (right) column, remove the
            // second item from the top without redistributing the survivors.
            if (rightPairs.length >= 2) {
                removePair(rightPairs[1]);
                orientationGapCreated = true;
            } else {
                shortEdgeWarning = true;
            }
        } else {
            var horizontalDistribution = calculateDistribution(
                productWidth,
                REGISTER_LONG,
                TARGET_ITEM_GAP
            );

            var topRegisterTop = productBounds[1] +
                PRODUCT_TO_REGISTER_GAP + REGISTER_SHORT;
            var bottomRegisterTop = productBounds[3] - PRODUCT_TO_REGISTER_GAP;
            var topPairs = [];
            var bottomPairs = [];
            var j;

            for (j = 0; j < horizontalDistribution.count; j++) {
                var registerLeft = distributionPosition(
                    productBounds[0],
                    productWidth,
                    REGISTER_LONG,
                    horizontalDistribution,
                    j,
                    false
                );

                var topPair = createPair(
                    registerLayer,
                    laserLayer,
                    topRegisterTop,
                    registerLeft,
                    REGISTER_LONG,
                    REGISTER_SHORT,
                    CUT_LONG,
                    CUT_SHORT,
                    registerColor,
                    laserColor,
                    "Top " + (j + 1),
                    createdItems
                );
                var bottomPair = createPair(
                    registerLayer,
                    laserLayer,
                    bottomRegisterTop,
                    registerLeft,
                    REGISTER_LONG,
                    REGISTER_SHORT,
                    CUT_LONG,
                    CUT_SHORT,
                    registerColor,
                    laserColor,
                    "Bottom " + (j + 1),
                    createdItems
                );

                topPairs.push(topPair);
                bottomPairs.push(bottomPair);
                allPairs.push(topPair);
                allPairs.push(bottomPair);
            }

            // Direction key: on the dot-adjacent (top) row, remove the second
            // item from the right without redistributing the survivors.
            if (topPairs.length >= 2) {
                removePair(topPairs[topPairs.length - 2]);
                orientationGapCreated = true;
            } else {
                shortEdgeWarning = true;
            }
        }

        // Normally the first and last Register items already span the selected
        // product bounds. Include the product explicitly for the one-item short
        // edge case so the resized artboard can never crop the finished path.
        var registerBounds = unionBounds(
            getActiveRegisterBounds(allPairs),
            productBounds
        );
        var newArtboardRect = [
            registerBounds[0] - ARTBOARD_MARGIN,
            registerBounds[1] + ARTBOARD_MARGIN,
            registerBounds[2] + ARTBOARD_MARGIN,
            registerBounds[3] - ARTBOARD_MARGIN
        ];

        var newArtboardWidth = newArtboardRect[2] - newArtboardRect[0];
        if (newArtboardWidth > 48 * PT_PER_INCH + EPSILON) {
            throw new Error(
                "The calculated artboard would be " + formatInches(newArtboardWidth) +
                " in wide, which exceeds 48 in."
            );
        }

        doc.artboards[activeArtboardIndex].artboardRect = newArtboardRect;
        changedArtboard = true;

        var orientationDot = laserLayer.pathItems.ellipse(
            newArtboardRect[1],
            newArtboardRect[2] - ORIENTATION_DOT_DIAMETER,
            ORIENTATION_DOT_DIAMETER,
            ORIENTATION_DOT_DIAMETER
        );
        styleLaserPath(orientationDot, laserColor, SMALL_CUT_STROKE);
        orientationDot.name = "Acrylic Orientation Dot";
        createdItems.push(orientationDot);

        // The original product paths are processed last, after all generated
        // geometry and the artboard have succeeded. Reparenting with move()
        // keeps their document coordinates unchanged.
        restyleAndMoveSelection(selection, laserLayer, laserColor);

        restoreLayerLocks(layerStates);
        app.redraw();

        if (shortEdgeWarning && !orientationGapCreated) {
            alert(
                "The edge was too short to create at least two small items per side. " +
                "The register items were created, but no direction gap was removed."
            );
        }
    } catch (error) {
        removeCreatedItems(createdItems);

        if (changedArtboard) {
            try {
                doc.artboards[activeArtboardIndex].artboardRect = oldArtboardRect;
            } catch (restoreArtboardError) {
                // Keep the original error as the useful message.
            }
        }

        restoreLayerStatesAfterFailure(layerStates);
        removeCreatedLayers(createdLayers);
        alert("Acrylic setup stopped:\n" + error.message);
    }

    function getSelectionBounds(items) {
        var bounds = null;
        var i;

        for (i = 0; i < items.length; i++) {
            var itemBounds;
            try {
                itemBounds = items[i].geometricBounds;
            } catch (error) {
                throw new Error("The selection contains an item with no usable geometric bounds.");
            }

            if (!itemBounds || itemBounds.length !== 4) {
                throw new Error("The selection contains an unsupported item type.");
            }

            bounds = unionBounds(bounds, itemBounds);
        }

        if (!bounds) {
            throw new Error("Could not calculate the selected artwork bounds.");
        }
        return bounds;
    }

    function validateSelectedLaserArtwork(items) {
        var pathCount = 0;
        var i;

        for (i = 0; i < items.length; i++) {
            pathCount += validateLaserArtworkItem(items[i]);
        }

        if (pathCount === 0) {
            throw new Error("The selection does not contain any editable laser-cut paths.");
        }
    }

    function validateLaserArtworkItem(item) {
        var count = 0;
        var i;

        if (item.typename === "PathItem") {
            return 1;
        }

        if (item.typename === "CompoundPathItem") {
            for (i = 0; i < item.pathItems.length; i++) {
                count += validateLaserArtworkItem(item.pathItems[i]);
            }
            return count;
        }

        if (item.typename === "GroupItem") {
            for (i = 0; i < item.pageItems.length; i++) {
                count += validateLaserArtworkItem(item.pageItems[i]);
            }
            return count;
        }

        throw new Error(
            "The original selection contains an unsupported item type: " + item.typename + "."
        );
    }

    function restyleAndMoveSelection(items, targetLayer, laserColor) {
        var i;

        for (i = 0; i < items.length; i++) {
            restyleLaserArtworkItem(items[i], laserColor);
        }

        for (i = 0; i < items.length; i++) {
            if (items[i].parent !== targetLayer) {
                items[i].move(targetLayer, ElementPlacement.PLACEATBEGINNING);
            }
        }
    }

    function restyleLaserArtworkItem(item, laserColor) {
        var i;

        if (item.typename === "PathItem") {
            styleLaserPath(item, laserColor, SMALL_CUT_STROKE);
            return;
        }

        if (item.typename === "CompoundPathItem") {
            for (i = 0; i < item.pathItems.length; i++) {
                restyleLaserArtworkItem(item.pathItems[i], laserColor);
            }
            return;
        }

        if (item.typename === "GroupItem") {
            for (i = 0; i < item.pageItems.length; i++) {
                restyleLaserArtworkItem(item.pageItems[i], laserColor);
            }
        }
    }

    function calculateDistribution(availableLength, itemLength, targetGap) {
        var maxFit = Math.floor((availableLength + EPSILON) / itemLength);

        if (maxFit < 2) {
            return {
                count: 1,
                gap: 0,
                centered: true
            };
        }

        var bestCount = 2;
        var bestGap = (availableLength - 2 * itemLength);
        var bestScore = Math.abs(bestGap - targetGap);
        var count;

        for (count = 3; count <= maxFit; count++) {
            var gap = (availableLength - count * itemLength) / (count - 1);
            var score = Math.abs(gap - targetGap);

            // On an exact tie, use more items.
            if (score <= bestScore + EPSILON) {
                bestCount = count;
                bestGap = gap;
                bestScore = score;
            }
        }

        return {
            count: bestCount,
            gap: bestGap,
            centered: false
        };
    }

    function distributionPosition(
        startCoordinate,
        availableLength,
        itemLength,
        distribution,
        index,
        vertical
    ) {
        if (distribution.centered) {
            if (vertical) {
                return startCoordinate - (availableLength - itemLength) / 2;
            }
            return startCoordinate + (availableLength - itemLength) / 2;
        }

        var step = itemLength + distribution.gap;
        if (vertical) {
            return startCoordinate - index * step;
        }
        return startCoordinate + index * step;
    }

    function createPair(
        registerLayer,
        laserLayer,
        registerTop,
        registerLeft,
        registerWidth,
        registerHeight,
        cutWidth,
        cutHeight,
        registerColor,
        laserColor,
        label,
        createdItems
    ) {
        var registerItem = registerLayer.pathItems.rectangle(
            registerTop,
            registerLeft,
            registerWidth,
            registerHeight
        );
        registerItem.stroked = false;
        registerItem.filled = true;
        registerItem.fillColor = registerColor;
        registerItem.fillOverprint = false;
        registerItem.name = "Acrylic Register " + label;
        createdItems.push(registerItem);

        var cutItem = laserLayer.pathItems.rectangle(
            registerTop - BLEED,
            registerLeft + BLEED,
            cutWidth,
            cutHeight
        );
        styleLaserPath(cutItem, laserColor, SMALL_CUT_STROKE);
        cutItem.name = "Acrylic Small Cut " + label;
        createdItems.push(cutItem);

        return {
            registerItem: registerItem,
            cutItem: cutItem,
            removed: false
        };
    }

    function styleLaserPath(pathItem, color, width) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeColor = color;
        pathItem.strokeWidth = width;
        pathItem.strokeOverprint = false;
    }

    function removePair(pair) {
        pair.registerItem.remove();
        pair.cutItem.remove();
        pair.removed = true;
    }

    function getActiveRegisterBounds(pairs) {
        var bounds = null;
        var i;

        for (i = 0; i < pairs.length; i++) {
            if (!pairs[i].removed) {
                bounds = unionBounds(bounds, pairs[i].registerItem.geometricBounds);
            }
        }

        if (!bounds) {
            throw new Error("No Register items were available for the final artboard bounds.");
        }
        return bounds;
    }

    function unionBounds(first, second) {
        if (!first) {
            return copyBounds(second);
        }

        return [
            Math.min(first[0], second[0]),
            Math.max(first[1], second[1]),
            Math.max(first[2], second[2]),
            Math.min(first[3], second[3])
        ];
    }

    function copyBounds(bounds) {
        return [bounds[0], bounds[1], bounds[2], bounds[3]];
    }

    function getOrCreateLayer(document, name, createdLayers) {
        var i;
        for (i = 0; i < document.layers.length; i++) {
            if (document.layers[i].name === name) {
                return document.layers[i];
            }
        }

        var layer = document.layers.add();
        layer.name = name;
        createdLayers.push(layer);
        return layer;
    }

    function prepareLayer(layer, states) {
        states.push({
            layer: layer,
            locked: layer.locked,
            visible: layer.visible
        });
        layer.locked = false;
        layer.visible = true;
    }

    function restoreLayerLocks(states) {
        var i;
        for (i = 0; i < states.length; i++) {
            try {
                states[i].layer.locked = states[i].locked;
            } catch (error) {
                // A completed setup is more important than restoring a UI lock.
            }
        }
    }

    function restoreLayerStatesAfterFailure(states) {
        var i;
        for (i = 0; i < states.length; i++) {
            try {
                states[i].layer.locked = states[i].locked;
                states[i].layer.visible = states[i].visible;
            } catch (error) {
                // Ignore cleanup-only failures.
            }
        }
    }

    function removeCreatedItems(items) {
        var i;
        for (i = items.length - 1; i >= 0; i--) {
            try {
                items[i].remove();
            } catch (error) {
                // The item may already have been removed as the direction gap.
            }
        }
    }

    function removeCreatedLayers(layers) {
        var i;
        for (i = layers.length - 1; i >= 0; i--) {
            try {
                if (layers[i].pageItems.length === 0) {
                    layers[i].remove();
                }
            } catch (error) {
                // Ignore cleanup-only failures.
            }
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

    function formatInches(points) {
        return (Math.round(points / PT_PER_INCH * 1000) / 1000).toString();
    }
}());
