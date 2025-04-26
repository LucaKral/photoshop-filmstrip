/**
 * Vertical Film Strip Generator for Photoshop 2025
 * Creates a filmstrip with dimensions scaled to 95% of original.
 * 3 Images, Duplicated, With Printing.
 * Keeps filmstrips black, aligns images, maintains aspect ratio, centers filmstrips and photos,
 * EVEN top/middle/bottom spacing, SELECT ALL 3 IMAGES AT ONCE, NO SPROCKET HOLES!, PRINT DIALOG, AUTO-SAVE AS "Filmstrip.jpg" IN IMAGE DIRECTORY, CLOSE WITHOUT SAVING, SPECIFIC DIMENSIONS, NO ALERT AT END.
 */

#target photoshop

// Function to open an image file
function openImage(filePath) {
    try {
        var file = new File(filePath);
        if (file.exists) {
            return open(file);
        } else {
            alert("Error: File not found - " + filePath);
            return null;
        }
    } catch (e) {
        alert("Error opening file: " + e);
        return null;
    }
}


// Function to resize an image proportionally
function resizeImage(doc, targetWidth, targetHeight) {
   doc.resizeImage(targetWidth, targetHeight, null, ResampleMethod.BICUBIC);
}

// Function to move all filmstrip layers to the bottom
function moveFilmstripLayersToBottom(doc) {
    var numLayers = doc.layers.length;
    for (var i = numLayers - 1; i >= 0; i--) {
        var layer = doc.layers[i];
        if (layer.name.indexOf("Filmstrip_") === 0) {
            layer.moveToEnd(doc);
        }
    }
}


// Main function
function main() {
    try {
        // *DEFINE ORIGINAL DOCUMENT DIMENSIONS*
        var originalDocWidthInches = 3.879;
        var originalDocHeightInches = 5.819;

        // *SCALE DIMENSIONS TO 95%*
        var docWidthInches = originalDocWidthInches * 0.95;
        var docHeightInches = originalDocHeightInches * 0.95;

        var resolution = 309; // PPI

        // Convert to pixels
        var docWidthPixels = docWidthInches * resolution;
        var docHeightPixels = docHeightInches * resolution;

        // Create new document
        var doc = app.documents.add(docWidthPixels, docHeightPixels, resolution, "FilmStrip", NewDocumentMode.RGB, DocumentFill.WHITE);

        // --- ADJUSTABLE PARAMETERS ---
        var numberOfFilmstrips = 2; // Number of filmstrips
        var filmstripWidthPercentage = 0.45; // Percentage of document width for each filmstrip
        var filmstripHeightPercentage = 0.90; // Percentage of document height for the filmstrip
        var filmstripSpacingPercentage = 0.05; // Percentage of document width for spacing between filmstrips
        var backgroundColor = new SolidColor();
        backgroundColor.rgb.red = 0;
        backgroundColor.rgb.green = 0;
        backgroundColor.rgb.blue = 0; // Black background

        // --- END ADJUSTABLE PARAMETERS ---

        // Filmstrip dimensions - VERTICAL
        var filmstripWidth = docWidthPixels * filmstripWidthPercentage;
        var filmstripHeight = docHeightPixels * filmstripHeightPercentage;

        var filmstripY = (docHeightPixels - filmstripHeight) / 2; // Center filmstrips vertically

        var filmstripX;
        var filmstripSpacing = docWidthPixels * filmstripSpacingPercentage;

       // Calculate the total width of the filmstrips and spacing
        var totalFilmstripWidth = (filmstripWidth * numberOfFilmstrips) + (filmstripSpacing * (numberOfFilmstrips - 1));

        // Calculate the starting X position to center the filmstrips
        var startX = (docWidthPixels - totalFilmstripWidth) / 2;

        // Create an array to store the filmstrip layers
        var filmstripLayers = [];

        // *GET INPUT FILES - SELECT 3 IMAGES AT ONCE*
        var filePaths = [];
        var files = File.openDialog("Select three images", "Images:*.jpg;*.jpeg;*.png;*.tif", true);
        if (files && files.length === 3) {
            filePaths = files;
        } else {
            alert("Please select exactly three images.");
            return;
        }

        // *DETERMINE SAVE DIRECTORY*
        var saveDirectory = new File(files[0]).parent; // Get directory from the first image

        // Loop through filmstrips
        for (var i = 0; i < numberOfFilmstrips; i++) {
            filmstripX = startX + (filmstripWidth + filmstripSpacing) * i; // Position horizontally

            // Create filmstrip background
            var filmstripLayer = doc.artLayers.add();
            filmstripLayer.name = "Filmstrip_" + (i + 1);
            filmstripLayers.push(filmstripLayer); // Add to the array

            // Draw black rectangle
            var filmstripRect = Array(
                [filmstripX, filmstripY],
                [filmstripX + filmstripWidth, filmstripY],
                [filmstripX + filmstripWidth, filmstripY + filmstripHeight],
                [filmstripX, filmstripY + filmstripHeight]
            );
            doc.selection.select(filmstripRect);

            doc.selection.fill(backgroundColor);
            doc.selection.deselect();

            // Calculate image frame dimensions
            var frameWidth = filmstripWidth;
            var frameHeight = filmstripHeight / 3;

           // *CORRECTED EVEN SPACING CALCULATION (FINALLY!)*
            var totalImageHeight = frameHeight * 3; // Total height of all images
            var totalSpacingHeight = filmstripHeight - totalImageHeight; // Total spacing height
            var evenSpacing = totalSpacingHeight / 4; // Divide the total spacing height by 4 to get even spacing

            // *LOOP THROUGH IMAGES FOR EACH FILMSTRIP*
            for (var j = 0; j < 3; j++) {


                // *CALCULATE THE CORRECT IMAGE INDEX*
                var imageIndex = j; // Use j directly as the index for the 3 images

                // If it's the second filmstrip, use the same images
                if (i === 1) {
                    imageIndex = j;
                }

                var imageX = filmstripX; // Position horizontally
                // *REVISED imageY CALCULATION*
                var imageY = filmstripY + (j * (frameHeight + evenSpacing)) + evenSpacing;

                // Open image
                var imageDoc = openImage(files[imageIndex]);
                if (imageDoc) {
                    // *RESIZE WHILE MAINTAINING ASPECT RATIO - CORRECTED*
                    var originalWidth = imageDoc.width;
                    var originalHeight = imageDoc.height;

                    var widthRatio = frameWidth / originalWidth;
                    var heightRatio = frameHeight / originalHeight;

                    var scaleRatio = Math.min(widthRatio, heightRatio);

                    var newWidth = originalWidth * scaleRatio;
                    var newHeight = originalHeight * scaleRatio;

                    // Resize the image
                    resizeImage(imageDoc, newWidth, newHeight);

                    // Copy and paste image
                    imageDoc.selection.selectAll();
                    imageDoc.selection.copy();
                    imageDoc.close(SaveOptions.DONOTSAVECHANGES);

                    // Paste the image
                    doc.paste();
                    var layerRef = doc.activeLayer;

                    // Center the image within the frame
                    layerRef.translate(imageX - layerRef.bounds[0] + (frameWidth - newWidth) / 2, imageY - layerRef.bounds[1] + (frameHeight - newHeight) / 2);
                }
            }
        }

        // *MOVE ALL FILMSTRIP LAYERS TO THE BOTTOM*
        moveFilmstripLayersToBottom(doc);

        // *OPEN PRINT DIALOG*
        app.displayDialogs = DialogModes.ALL; // Ensure dialogs are displayed
        doc.print();
        app.displayDialogs = DialogModes.ERROR; // Restore default error-only dialogs

        // *AUTO-SAVE AS JPG*
        var jpgSaveOptions = new JPEGSaveOptions();
        jpgSaveOptions.quality = 10;
        var savePath = new File(saveDirectory + "/Filmstrip.jpg"); // Construct the save path
        doc.saveAs(savePath, jpgSaveOptions, true, Extension.LOWERCASE);

        // *CLOSE WITHOUT SAVING*
        doc.close(SaveOptions.DONOTSAVECHANGES);


    } catch (e) {
        alert("Script Error: " + e);
    }
}

main();