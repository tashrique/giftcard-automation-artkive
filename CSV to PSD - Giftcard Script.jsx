// Initialize Files
{
    var rootFolder = File($.fileName).parent;
    var logoFile = rootFolder + "/Artkive_GC.png";
    var backgroundFile = rootFolder + "/Giftcards.png";

    // var fileName = File.openDialog("Select a file ~ CSV ~ |To:;|Message;|From;|OrderID;|");
    // var csvFile = new File(fileName);
    // var outputFolder = Folder.selectDialog("Select a folder to save Output"); // Output directory

    //Filepath for testing purposes
    var fileName = File(rootFolder + "\\Output\\formatted-csv.csv"); //CSV file
    var outputFolder = rootFolder + "\\Output"; // Output directory
    var csvFile = new File(fileName);

    //Get time for subfolder and file naming
    var date = new Date();
    var dateString = ('0' + (date.getMonth() + 1)).slice(-2) + '-' +
        ('0' + date.getDate()).slice(-2) + '-' +
        date.getFullYear() + ' ';
    var timeString = ((date.getHours())) + 'H-' +
        (date.getMinutes()) + 'M-' +
        date.getSeconds() + 'S';

    //Create Subfolder
    var subFolder = new Folder(outputFolder + '\\Batch - ' + dateString + timeString);
    if (!subFolder.exists) { subFolder.create() };


    //Check if selected folder is valid
    if (subFolder == null) {
        alert("Invalid folder. Please select a document and restart script.");
    }
}
//Check if selected csv file is valid
if (csvFile != null) {

    //Some maths for CSV file
    {
        csvFile.open('r');
        var csvData = csvFile.read();
        var lines = csvData.split('\n');
        csvFile.close();
        var numRows = lines.length - 1;
        var numDocs = Math.floor(numRows / 4);
        var remainder = numRows % 4;

        if (remainder != 0) {
            numDocs++;
        }
    }


    //Build Template
    {
        //Create Document
        var docRef = app.documents
        docRef.add(UnitValue(2550, "px"), UnitValue(3300, "px"), 300), "Giftcards-1"; // pixels

        var recRad = 100;

        //Draw Box, To, From and Message Shapes function
        // {

        //     //Helper Functions for Shape Drawing
        //     {
        //         function drawShapeHelper(R, G, B) {
        //             var desc88 = new ActionDescriptor();
        //             var ref60 = new ActionReference();
        //             ref60.putClass(stringIDToTypeID("contentLayer"));
        //             desc88.putReference(charIDToTypeID("null"), ref60);
        //             var desc89 = new ActionDescriptor();
        //             var desc90 = new ActionDescriptor();
        //             var desc91 = new ActionDescriptor();
        //             desc91.putDouble(charIDToTypeID("Rd  "), R); // R
        //             desc91.putDouble(charIDToTypeID("Grn "), G); // G
        //             desc91.putDouble(charIDToTypeID("Bl  "), B); // B
        //             var id481 = charIDToTypeID("RGBC");
        //             desc90.putObject(charIDToTypeID("Clr "), id481, desc91);
        //             desc89.putObject(charIDToTypeID("Type"), stringIDToTypeID("solidColorLayer"), desc90);
        //             desc88.putObject(charIDToTypeID("Usng"), stringIDToTypeID("contentLayer"), desc89);
        //             executeAction(charIDToTypeID("Mk  "), desc88, DialogModes.NO);
        //         }
        //     }


        //     {
        //         function DrawShapeTL() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(248, 240, 221); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }

        //     {
        //         function DrawShapeTR() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(226, 235, 230); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }

        //     {
        //         function DrawShape() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(226, 235, 230); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }



        //     {
        //         function DrawShapeBL() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(250, 231, 235); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }



        //     {
        //         function DrawShapeBR() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(226, 236, 236); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }


        //     // All TO Lines
        //     {
        //         function DrawShapeTLToLine() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(255, 255, 255); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }

        //     {
        //         function DrawShapeTRToLine() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(255, 255, 255); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }


        //     {
        //         function DrawShapeBLToLine() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(255, 255, 255); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }


        //     {
        //         function DrawShapeBRToLine() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(255, 255, 255); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }

        //     // All FROM Lines
        //     {
        //         function DrawShapeTLFromLine() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(255, 255, 255); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }

        //     {
        //         function DrawShapeTRFromLine() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(255, 255, 255); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }


        //     {
        //         function DrawShapeBLFromLine() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(255, 255, 255); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }


        //     {
        //         function DrawShapeBRFromLine() {

        //             var doc = app.activeDocument;
        //             var y = arguments.length;
        //             var i = 0;

        //             var lineArray = [];
        //             for (i = 0; i < y; i++) {
        //                 lineArray[i] = new PathPointInfo;
        //                 lineArray[i].kind = PointKind.CORNERPOINT;
        //                 lineArray[i].anchor = arguments[i];
        //                 lineArray[i].leftDirection = lineArray[i].anchor;
        //                 lineArray[i].rightDirection = lineArray[i].anchor;
        //             }

        //             var lineSubPathArray = new SubPathInfo();
        //             lineSubPathArray.closed = true;
        //             lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //             lineSubPathArray.entireSubPath = lineArray;
        //             var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //             drawShapeHelper(255, 255, 255); //R, G, B

        //             myPathItem.remove();
        //         }
        //     }


        //     //All MESSAGE Line
        //     {
        //         {
        //             function DrawShapeTLMesgLine() {

        //                 var doc = app.activeDocument;
        //                 var y = arguments.length;
        //                 var i = 0;

        //                 var lineArray = [];
        //                 for (i = 0; i < y; i++) {
        //                     lineArray[i] = new PathPointInfo;
        //                     lineArray[i].kind = PointKind.CORNERPOINT;
        //                     lineArray[i].anchor = arguments[i];
        //                     lineArray[i].leftDirection = lineArray[i].anchor;
        //                     lineArray[i].rightDirection = lineArray[i].anchor;
        //                 }

        //                 var lineSubPathArray = new SubPathInfo();
        //                 lineSubPathArray.closed = true;
        //                 lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //                 lineSubPathArray.entireSubPath = lineArray;
        //                 var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //                 drawShapeHelper(255, 255, 255); //R, G, B

        //                 myPathItem.remove();
        //             }
        //         }


        //         {
        //             function DrawShapeTRMesgLine() {

        //                 var doc = app.activeDocument;
        //                 var y = arguments.length;
        //                 var i = 0;

        //                 var lineArray = [];
        //                 for (i = 0; i < y; i++) {
        //                     lineArray[i] = new PathPointInfo;
        //                     lineArray[i].kind = PointKind.CORNERPOINT;
        //                     lineArray[i].anchor = arguments[i];
        //                     lineArray[i].leftDirection = lineArray[i].anchor;
        //                     lineArray[i].rightDirection = lineArray[i].anchor;
        //                 }

        //                 var lineSubPathArray = new SubPathInfo();
        //                 lineSubPathArray.closed = true;
        //                 lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //                 lineSubPathArray.entireSubPath = lineArray;
        //                 var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //                 drawShapeHelper(255, 255, 255); //R, G, B

        //                 myPathItem.remove();
        //             }
        //         }



        //         {
        //             function DrawShapeBLMesgLine() {

        //                 var doc = app.activeDocument;
        //                 var y = arguments.length;
        //                 var i = 0;

        //                 var lineArray = [];
        //                 for (i = 0; i < y; i++) {
        //                     lineArray[i] = new PathPointInfo;
        //                     lineArray[i].kind = PointKind.CORNERPOINT;
        //                     lineArray[i].anchor = arguments[i];
        //                     lineArray[i].leftDirection = lineArray[i].anchor;
        //                     lineArray[i].rightDirection = lineArray[i].anchor;
        //                 }

        //                 var lineSubPathArray = new SubPathInfo();
        //                 lineSubPathArray.closed = true;
        //                 lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //                 lineSubPathArray.entireSubPath = lineArray;
        //                 var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //                 drawShapeHelper(255, 255, 255); //R, G, B

        //                 myPathItem.remove();
        //             }
        //         }


        //         {
        //             function DrawShapeBLMesgLine() {

        //                 var doc = app.activeDocument;
        //                 var y = arguments.length;
        //                 var i = 0;

        //                 var lineArray = [];
        //                 for (i = 0; i < y; i++) {
        //                     lineArray[i] = new PathPointInfo;
        //                     lineArray[i].kind = PointKind.CORNERPOINT;
        //                     lineArray[i].anchor = arguments[i];
        //                     lineArray[i].leftDirection = lineArray[i].anchor;
        //                     lineArray[i].rightDirection = lineArray[i].anchor;
        //                 }

        //                 var lineSubPathArray = new SubPathInfo();
        //                 lineSubPathArray.closed = true;
        //                 lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //                 lineSubPathArray.entireSubPath = lineArray;
        //                 var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //                 drawShapeHelper(255, 255, 255); //R, G, B

        //                 myPathItem.remove();
        //             }
        //         }


        //         {
        //             function DrawShapeBRMesgLine() {

        //                 var doc = app.activeDocument;
        //                 var y = arguments.length;
        //                 var i = 0;

        //                 var lineArray = [];
        //                 for (i = 0; i < y; i++) {
        //                     lineArray[i] = new PathPointInfo;
        //                     lineArray[i].kind = PointKind.CORNERPOINT;
        //                     lineArray[i].anchor = arguments[i];
        //                     lineArray[i].leftDirection = lineArray[i].anchor;
        //                     lineArray[i].rightDirection = lineArray[i].anchor;
        //                 }

        //                 var lineSubPathArray = new SubPathInfo();
        //                 lineSubPathArray.closed = true;
        //                 lineSubPathArray.operation = ShapeOperation.SHAPEADD;
        //                 lineSubPathArray.entireSubPath = lineArray;
        //                 var myPathItem = doc.pathItems.add("myPath", [lineSubPathArray]);

        //                 drawShapeHelper(255, 255, 255); //R, G, B

        //                 myPathItem.remove();
        //             }
        //         }
        //     }
        // }

        // //Place Logo Function
        // {
        //     function placeLogo(X, Y) {
        //         var originalDoc = app.activeDocument;
        //         var fileRef = File(logoFile);
        //         var doc = app.open(fileRef);


        //         // duplicate to original document
        //         var layer = doc.activeLayer;
        //         var newLayer = layer.duplicate(originalDoc, ElementPlacement.PLACEATBEGINNING);

        //         // close new document
        //         doc.close(SaveOptions.DONOTSAVECHANGES);

        //         var doc = app.activeDocument;
        //         var layer = doc.activeLayer;
        //         layer.translate(X, Y);
        //     }
        // }

        // //To and From Lines Function
        // {
        //     function writeToFrom(Text, X, Y, R, G, B) {
        //         var doc = app.activeDocument;

        //         var layerTO = doc.artLayers.add();
        //         layerTO.kind = LayerKind.TEXT;
        //         layerTO.textItem.contents = Text;
        //         layerTO.textItem.position = [X, Y]; // set the position in picas
        //         layerTO.textItem.size = 12; // set the font size in points
        //         layerTO.textItem.font = "Montserrat-Bold";
        //         layerTO.textItem.color.rgb.red = R;
        //         layerTO.textItem.color.rgb.green = G;
        //         layerTO.textItem.color.rgb.blue = B;
        //     }
        // }

        // {
        //     function templateTL() {
        //         DrawShapeTL([18, 18], [288, 18], [288, 378], [18, 378]);
        //         DrawShapeTLToLine([36, 36], [270, 36], [270, 57], [36, 57]);
        //         DrawShapeTLFromLine([36, 61], [270, 61], [270, 82], [36, 82]);
        //         DrawShapeTLMesgLine([36, 96], [270, 96], [270, 277], [36, 277]);
        //         activeDocument.mergeVisibleLayers();
        //         placeLogo(40, 292);
        //         writeToFrom("To", 44, 51, 248, 240, 221);
        //         writeToFrom("From", 44, 75, 248, 240, 221);

        //     }

        //     function templateTR() {
        //         DrawShapeTR([325, 18], [591, 18], [591, 379], [325, 379]);
        //         DrawShapeTRToLine([342, 36], [577, 36], [577, 57], [342, 57]);
        //         DrawShapeTRFromLine([342, 61], [577, 61], [577, 82], [342, 82]);
        //         DrawShapeTRMesgLine([342, 96], [577, 96], [577, 277], [342, 277]);
        //         activeDocument.mergeVisibleLayers();
        //         placeLogo(341, 292);
        //         writeToFrom("To", 351, 51, 226, 235, 230);
        //         writeToFrom("From", 351, 75, 226, 235, 230);

        //     }

        //     function templateBL() {
        //         DrawShapeBL([18, 415], [288, 415], [288, 772], [18, 772]);
        //         DrawShapeBLToLine([36, 432], [270, 432], [270, 452], [36, 452]);
        //         DrawShapeBLFromLine([36, 457], [270, 457], [270, 477], [36, 477]);
        //         DrawShapeBLMesgLine([36, 492], [270, 490], [270, 672], [36, 672]);
        //         activeDocument.mergeVisibleLayers();
        //         placeLogo(40, 688);
        //         writeToFrom("To", 44, 447, 250, 231, 235);
        //         writeToFrom("From", 44, 470, 250, 231, 235);

        //     }

        //     function templateBR() {
        //         DrawShapeBR([325, 415], [594, 415], [594, 772], [325, 772]);
        //         DrawShapeBRToLine([342, 432], [577, 432], [577, 452], [342, 452]);
        //         DrawShapeBRFromLine([342, 457], [577, 457], [577, 477], [342, 477]);
        //         DrawShapeBRMesgLine([342, 492], [577, 492], [577, 672], [342, 672]);
        //         activeDocument.mergeVisibleLayers();
        //         placeLogo(341, 688);
        //         writeToFrom("To", 351, 447, 226, 235, 230);
        //         writeToFrom("From", 351, 470, 226, 235, 230);

        //     }

        //     function template() {
        //         templateTL();
        //         templateTR();
        //         templateBL();
        //         templateBR();
        //         activeDocument.mergeVisibleLayers();
        //     }
        // }

        //MAKEEE THE TEMPLATEEE AAAAAAAAAAAAAAAA!
        //template();

        //Place Background Function
        {
            function placeBackground() {
                var originalDoc = app.activeDocument;
                var fileRef = File(backgroundFile);
                var doc = app.open(fileRef);


                // duplicate to original document
                var layer = doc.activeLayer;
                var newLayer = layer.duplicate(originalDoc, ElementPlacement.PLACEATBEGINNING);

                // close new document
                doc.close(SaveOptions.DONOTSAVECHANGES);

                var doc = app.activeDocument;
                var layer = doc.activeLayer;
                layer.translate(0, 0);
                activeDocument.mergeVisibleLayers();

            }
        }
        placeBackground();


    }

    /*------------------Template Built At this Point LESSGOOOO-----------------------*/




    // Loop through all lines and split them into cells (storing them in and array)
    var data = [];
    for (var i = 0; i < numRows; i++) {
        data.push(lines[i].split(','));
    }

    //Read CSV File and Split into usable format
    function addText(p) {
        if (documents.length > 0) {
            var originalDialogMode = app.displayDialogs;
            app.displayDialogs = DialogModes.ERROR;
            var originalRulerUnits = preferences.rulerUnits;
            preferences.rulerUnits = Units.PIXELS;
            var docRef = app.activeDocument;


            for (var i = (p * 4); i < lines.length && i < ((p * 4) + 4); i++) {
                var cell = data[i]; // split each line into cells
                var LayerRef = null;
                var TextRef = null;

                if (cell == undefined || !cell || typeof (cell) != "object" || cell == null || cell == "undefined") {
                    continue;
                }

                // create a new text layer and set its contents to the cell value
                LayerRef = docRef.artLayers.add();
                LayerRef.kind = LayerKind.TEXT;
                TextRef = LayerRef.textItem;
                TextRef.contents = cell.join(',');

                // Optional Styling
                {
                    preferences.rulerUnits = Units.POINTS;
                    TextRef.size = 24;
                    TextRef.useAutoLeading = false;
                    TextRef.leading = 24;
                    TextRef.font = 'Futura-PT';
                    TextRef.justification = Justification.CENTER;
                    TextRef.autoKerning = AutoKernType.METRICS;
                    // EOF text styling
                }
            }
            csvFile.close();

            preferences.rulerUnits = originalRulerUnits;
            app.displayDialogs = originalDialogMode;

        }
        else {
            alert('You must have a document open to run this script.');
            return;
        }
        return;
    }


    for (var p = 0; p < numDocs; p++) {

        // User input placement
        addText(p);
        var doc = app.activeDocument;
        var selectedLayer = doc.activeLayer;

        var remaining = numRows - (p * 4);

        //Global function to generate and format user questions
        function userToFromMessageID(toX, toY, fromX, fromY, bodyX, bodyY, IDX, IDY, variable) {

            //alert("Variable array: " + variable);
            if (!variable || variable == undefined || variable == "undefined") {
                variable = "    ";
            }

            //TO line
            var doc = app.activeDocument;
            var layerTL1 = doc.artLayers.add();
            layerTL1.kind = LayerKind.TEXT;
            var tempo = variable[0].slice(1);
            layerTL1.textItem.contents = tempo;
            layerTL1.textItem.position = [toX, toY]; // set the position in picas
            layerTL1.textItem.size = 11; // set the font size in points
            layerTL1.textItem.font = "Futura-PT";
            layerTL1.textItem.color.rgb.red = 0;
            layerTL1.textItem.color.rgb.green = 0;
            layerTL1.textItem.color.rgb.blue = 0;

            //FROM line
            var doc = app.activeDocument;
            var layerTL2 = doc.artLayers.add();
            layerTL2.kind = LayerKind.TEXT;
            layerTL2.textItem.contents = variable[2];
            layerTL2.textItem.position = [fromX, fromY]; // set the position in picas
            layerTL2.textItem.size = 11; // set the font size in points
            layerTL2.textItem.font = "Futura-PT";
            layerTL2.textItem.color.rgb.red = 0;
            layerTL2.textItem.color.rgb.green = 0;
            layerTL2.textItem.color.rgb.blue = 0;

            //Message Body
            var doc = app.activeDocument;
            var layerTL3 = doc.artLayers.add();
            layerTL3.kind = LayerKind.TEXT;
            layerTL3.textItem.contents = variable[1];
            layerTL3.textItem.position = [bodyX, bodyY]; // set the position in picas
            layerTL3.textItem.size = 11; // set the font size in points
            layerTL3.textItem.font = "Futura-PT";
            layerTL3.textItem.color.rgb.red = 0;
            layerTL3.textItem.color.rgb.green = 0;
            layerTL3.textItem.color.rgb.blue = 0;
            // convert the text layer to a paragraph text layer
            layerTL3.textItem.kind = TextType.PARAGRAPHTEXT;
            // set the width and height of the text box
            layerTL3.textItem.width = 220; // set the width in pixels
            layerTL3.textItem.height = 178; // set the height in pixels

            //Order ID
            var doc = app.activeDocument;
            var layerTL4 = doc.artLayers.add();
            layerTL4.kind = LayerKind.TEXT;
            layerTL4.textItem.contents = variable[3];
            layerTL4.textItem.position = [IDX, IDY]; // set the position in picas
            layerTL4.textItem.size = 11; // set the font size in points
            layerTL3.textItem.font = "Futura-PT";
            layerTL4.textItem.color.rgb.red = 255;
            layerTL4.textItem.color.rgb.green = 255;
            layerTL4.textItem.color.rgb.blue = 255;
        }


        if (remainder == 0 || (remaining != 1 && remaining != 2 && remaining != 3)) {
            //Split the row inputs into different fields
            {
                var textLayer1 = doc.layers[3];
                var contents = textLayer1.textItem.contents;
                var variables1 = contents.split(';');

                var textLayer2 = doc.layers[2];
                var contents2 = textLayer2.textItem.contents;
                var variables2 = contents2.split(';');

                var textLayer3 = doc.layers[1];
                var contents3 = textLayer3.textItem.contents;
                var variables3 = contents3.split(';');

                var textLayer4 = doc.layers[0];
                var contents4 = textLayer4.textItem.contents;
                var variables4 = contents4.split(';');
            }

            //Hide the original text layers
            {
                var doc = app.activeDocument;
                var layers = doc.layers;

                for (var i = 0; i < layers.length; i++) {
                    var layer = layers[i];
                    if (layer.kind == LayerKind.TEXT) {
                        layer.visible = false;
                    }
                }
            }

            //Format user inputs
            {
                //Arguments: toX, toY, fromX, fromY, bodyX, bodyY, IDX, IDY

                //TL
                userToFromMessageID(86, 51, 86, 75, 44, 115, 233, 373, variables1);

                //TR
                userToFromMessageID(392, 51, 392, 75, 350, 115, 538, 373, variables2);

                //BL
                userToFromMessageID(86, 447, 86, 470, 44, 511, 233, 768, variables3);

                //BR
                userToFromMessageID(392, 447, 392, 470, 350, 511, 538, 768, variables4);

            }

        }

        if (remaining == 1) {

            //Split the row inputs into different fields
            {
                var textLayer1 = doc.layers[0];
                var contents = textLayer1.textItem.contents || " ";
                var variables1 = contents.split(';');
            }

            //Hide the original text layers
            {
                var doc = app.activeDocument;
                var layers = doc.layers;

                for (var i = 0; i < layers.length; i++) {
                    var layer = layers[i];
                    if (layer.kind == LayerKind.TEXT) {
                        layer.visible = false;
                    }
                }
            }

            //Format user inputs
            {
                //Arguments: toX, toY, fromX, fromY, bodyX, bodyY, IDX, IDY
                //TL
                userToFromMessageID(86, 51, 86, 75, 44, 115, 233, 373, variables1);
            }
        }

        if (remaining == 2) {

            //Split the row inputs into different fields
            {
                var textLayer1 = doc.layers[0];
                var contents1 = textLayer1.textItem.contents || " ";
                var variables1 = contents1.split(';');

                var textLayer2 = doc.layers[1];
                var contents2 = textLayer2.textItem.contents || " ";
                var variables2 = contents2.split(';');
            }

            //Hide the original text layers
            {
                var doc = app.activeDocument;
                var layers = doc.layers;

                for (var i = 0; i < layers.length; i++) {
                    var layer = layers[i];
                    if (layer.kind == LayerKind.TEXT) {
                        layer.visible = false;
                    }
                }
            }

            //Format user inputs
            {
                //Arguments: toX, toY, fromX, fromY, bodyX, bodyY, IDX, IDY
                //TL
                userToFromMessageID(86, 51, 86, 75, 44, 115, 233, 373, variables1);

                //TR
                userToFromMessageID(392, 51, 392, 75, 350, 115, 538, 373, variables2);
            }
        }

        if (remaining == 3) {

            //Split the row inputs into different fields
            {
                var textLayer1 = doc.layers[0];
                var contents1 = textLayer1.textItem.contents || " ";
                var variables1 = contents1.split(';');

                var textLayer2 = doc.layers[1];
                var contents2 = textLayer2.textItem.contents || " ";
                var variables2 = contents2.split(';');

                var textLayer3 = doc.layers[2];
                var contents3 = textLayer3.textItem.contents || " ";
                var variables3 = contents3.split(';');
            }

            //Hide the original text layers
            {
                var doc = app.activeDocument;
                var layers = doc.layers;

                for (var i = 0; i < layers.length; i++) {
                    var layer = layers[i];
                    if (layer.kind == LayerKind.TEXT) {
                        layer.visible = false;
                    }
                }
            }

            //Format user inputs
            {
                //Arguments: toX, toY, fromX, fromY, bodyX, bodyY, IDX, IDY
                //TL
                userToFromMessageID(86, 51, 86, 75, 44, 115, 233, 373, variables1);

                //TR
                userToFromMessageID(392, 51, 392, 75, 350, 115, 538, 373, variables2);

                //BL
                userToFromMessageID(86, 447, 86, 470, 44, 511, 233, 768, variables3);
            }
        }

        //Save Doc as PNG in format giftcards+MMDDYYY.pdf
        {

            //var doc = app.activeDocument;


            var fileName = "giftcards" + dateString + "- " + p + ".png";
            var outputFile = new File(subFolder + '/' + fileName);

            // Save as PNG
            var pngSaveOptions = new PNGSaveOptions();
            pngSaveOptions.interlaced = true;
            doc.saveAs(outputFile, pngSaveOptions, true);

            // Save as PSD
            var psdSaveOptions = new PhotoshopSaveOptions();
            psdSaveOptions.embedColorProfile = true;
            psdSaveOptions.alphaChannels = true;
            psdSaveOptions.layers = true;  // Preserve layers
            doc.saveAs(outputFile, psdSaveOptions, true);
        }

        //alert((p + 1) + "-th iteration. Click OK for next 4 cards.");

        //Delete all user input layers
        {
            // Check if doc is open
            if (app.documents.length > 0) {
                var doc = app.activeDocument;

                // loop backwards through the layers because deleting a layer affects the collection indices
                for (var i = doc.layers.length - 1; i >= 0; i--) { // Changed the condition to i >= 0
                    var layer = doc.layers[i];

                    // Check if the layer is the background layer
                    if (layer.isBackgroundLayer) {
                        // If it's the background layer, skip the current iteration
                        continue;
                    }
                    layer.remove();
                }
            }
        }
    }

    //Close Doc if all scripts successful
    // alert("Script Complete. Created " + p + " sheet(s) of Giftcards.");
    var doc = app.activeDocument;
    doc.close(SaveOptions.DONOTSAVECHANGES);
}

else {
    alert("Invalid File. Please select a CSV file and restart script")
}

