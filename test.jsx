#target photoshop

var doc = app.activeDocument;
var layer = doc.activeLayer;

// Load transparency as selection (selects opaque pixels)
var desc = new ActionDescriptor();
var ref = new ActionReference();
ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
desc.putReference(charIDToTypeID("null"), ref);
var ref2 = new ActionReference();
ref2.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Trsp"));
desc.putReference(charIDToTypeID("T   "), ref2);
executeAction(charIDToTypeID("setd"), desc, DialogModes.NO);

// Invert selection to get non-opaque pixels, then delete them
doc.selection.invert();
doc.selection.clear();
doc.selection.deselect();

// Show the bounds of the cleaned up layer
var b = layer.bounds;
alert(
    "Bounds for layer: " + layer.name + "\n\n" +
    "left: " + b[0].as("px") + "\n" +
    "top: " + b[1].as("px") + "\n" +
    "right: " + b[2].as("px") + "\n" +
    "bottom: " + b[3].as("px") + "\n\n" +
    "width: " + (b[2].as("px") - b[0].as("px")) + "\n" +
    "height: " + (b[3].as("px") - b[1].as("px"))
);
