/**
 * Code 128 Barcode Generator for Adobe Illustrator
 * Generates text and numbers
 * Created by Katja G. Bjerrum
 */


var CODE_128B_PATTERNS = {
    startCode: "11010010000", // Start Code B (104)
    stopCode: "1100011101011", // Stop Code (106)
    patterns: [
        "11011001100", "11001101100", "11001100110", "10010011000",
        "10010001100", "10001001100", "10011001000", "10011000100",
        "10001100100", "11001001000", "11001000100", "11000100100",
        "10110011100", "10011011100", "10011001110", "10111001100",
        "10011101100", "10011100110", "11001110010", "11001011100",
        "11001001110", "11011100100", "11001110100", "11101101110",
        "11101001100", "11100101100", "11100100110", "11101100100",
        "11100110100", "11100110010", "11011011000", "11011000110",
        "11000110110", "10100011000", "10001011000", "10001000110",
        "10110001000", "10001101000", "10001100010", "11010001000",
        "11000101000", "11000100010", "10110111000", "10110001110",
        "10001101110", "10111011000", "10111000110", "10001110110",
        "11101110110", "11010001110", "11000101110", "11011101000",
        "11011100010", "11011101110", "11101011000", "11101000110",
        "11100010110", "11101101000", "11101100010", "11100011010",
        "11101111010", "11001000010", "11110001010", "10100110000",
        "10100001100", "10010110000", "10010000110", "10000101100",
        "10000100110", "10110010000", "10110000100", "10011010000",
        "10011000010", "10000110100", "10000110010", "11000010010",
        "11001010000", "11110111010", "11000010100", "10001111010",
        "10100111100", "10010111100", "10010011110", "10111100100",
        "10011110100", "10011110010", "11110100100", "11110010100",
        "11110010010", "11011011110", "11011110110", "11110110110",
        "10101111000", "10100011110", "10001011110", "10111101000",
        "10111100010", "11110101000", "11110100010", "10111011110",
        "10111101110", "11101011110", "11110101110"
    ]
};

function createCode128BarcodesForSelection() {
    if (app.documents.length === 0) {
        alert("Please open a document first!");
        return;
    }

    var doc = app.activeDocument;

    if (doc.selection.length === 0) {
        alert("Please select one or more text objects!");
        return;
    }

    for (var i = 0; i < doc.selection.length; i++) {
        var selectedItem = doc.selection[i];
        if (selectedItem.typename === "TextFrame" || selectedItem.typename === "PointText") {
            createBarcodeForTextItem(selectedItem);
        } else {
            alert("Item " + (i + 1) + " is not a text object. Skipping.");
        }
    }

    
}

function createBarcodeForTextItem(textItem) {
    var message = textItem.contents;
    if (message.length === 0) {
        alert("Text frame is empty. Skipping.");
        return;
    }
    var matrix = textItem.matrix;
    var angle = -180 / Math.PI * Math.atan2(matrix.mValueC, matrix.mValueD);
    
if(angle !== 0){
    
        var tempText = textItem.duplicate();
    tempText.rotate(-angle);

    var textBounds = tempText.geometricBounds;
    tempText.remove(); // ✅ Clean up the temporary text


}
else{
 var textBounds = textItem.geometricBounds;

}
 
    var textWidth = textBounds[2] - textBounds[0];
    var textHeight = Math.abs(textBounds[1] - textBounds[3]);

    // Create a group to hold the text and barcode
    var group = app.activeDocument.groupItems.add();

    // Add the text item to the group
    textItem.move(group, ElementPlacement.PLACEATEND);

    // Generate pattern
    var pattern = generateCode128Pattern(message);

    // Calculate dimensions
    // Target barcode width should be 20% wider than text width
    var targetBarcodeWidth = textWidth * 1.2;
    
    // Calculate bar width based on the desired total width
    var barWidth = targetBarcodeWidth / pattern.length;
    
    var barcodeHeight = textHeight * 4.5; // Barcode height
    var barcodeWidth = pattern.length * barWidth;

    // Calculate starting X position to center the barcode over the text
    // This ensures the barcode is centered regardless of text justification
    var startX = textBounds[0] + (textWidth - barcodeWidth) / 2;
    var currentX = startX;

    // Create a group for the barcode
    var barcodeGroup = group.groupItems.add();

    // Create black color
    var blackColor = new CMYKColor();
    blackColor.black = 100;

    // Draw each bar
    for (var i = 0; i < pattern.length; i++) {
        if (pattern[i] === "1") {
            var bar = barcodeGroup.pathItems.rectangle(
                textBounds[1] + 3, // Position above the text with a smaller gap
                currentX,
                barWidth,
                -barcodeHeight
            );
            bar.filled = true;
            bar.fillColor = blackColor;
            bar.stroked = false;
        }
        currentX += barWidth;
    }
    // Determine rotation from the text's transformation matrix
    var matrix = textItem.matrix;
    var angle = -180 / Math.PI * Math.atan2(matrix.mValueC, matrix.mValueD);

    // Calculate the rotation center
    var rotationX = textItem.left + textItem.width / 2;
    var rotationY = textItem.top - textItem.height / 2;

    // Apply rotation to the barcode group
    rotate_around_point(barcodeGroup, rotationY, rotationX, angle);

}

function generateCode128Pattern(message) {
    var pattern = CODE_128B_PATTERNS.startCode;
    var checksum = 104; // Start B value

    for (var i = 0; i < message.length; i++) {
        var charCode = message.charCodeAt(i) - 32;
        if (charCode < 0 || charCode >= CODE_128B_PATTERNS.patterns.length) {
            alert("Invalid character in input: " + message[i]);
            return "";
        }
        pattern += CODE_128B_PATTERNS.patterns[charCode];
        checksum += (i + 1) * charCode;
    }

    checksum = checksum % 103;
    pattern += CODE_128B_PATTERNS.patterns[checksum];
    pattern += CODE_128B_PATTERNS.stopCode;

    return pattern;
}

// Run the script
try {
    createCode128BarcodesForSelection();
} catch (e) {
    alert("Error creating barcodes: " + e.message);
}

function rotate_around_point(obj, roto_X, roto_Y, angle) {
    // Utility that rotates an object around an arbitrary point

    //Translate from point to document origin 
    obj.translate(-roto_Y, -roto_X);

    //Rotate around document origin
    obj.rotate(angle, true, true, true, true, Transformation.DOCUMENTORIGIN);

    //Translate back
    obj.translate(roto_Y, roto_X);
};