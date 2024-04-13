//*******************************************************
// Corius_Editable_Barcodes.jsx
// Version 1.0
//
// Copyright 2024 Corius
// Comments or suggestions to contact@corius.fr
//
//*******************************************************

// Script version
var Corius_Editable_Barcodes_version = 'v1.0';

// CONSOLE BLANK LINE
$.writeln('          ----------<<<<<<<<<<     EXEC START     >>>>>>>>>>----------');

// AI document variables
var docObj = app.activeDocument;
var docName = docObj.fullName;
var docFolder = docName.parent.fsName;

var barcodeObjectList = new Array();
var barcodeToEditList = new Array();
var CoriusBarcodeNames = new Array('CoriusBarcodeEAN13','CoriusBarcodeITF14');
var CSVReportTxt = 'Full code submitted;Check digit from code submitted;Corrected check digit;Corrected code'+'\n';

var correctedErrorList = new Array();

// document settings relevant
var unitTable = new Array();
unitTable[0] = new Array('Full text', 'Unknown', 'Inches', 'Centimeters', 'Points', 'Picas', 'Millimeters', 'Qs', 'Pixels', 'FeetInches', 'Meters', 'Yards', 'Feet');
unitTable[1] = new Array('mm equivalent', 'Unknown', 25.4, 10, null, null, 1, null, 25.4/72, null, 1000, 914.4, 304.8);
var unitRatio = null;
var svgRatio = 25.4/72;

//////////////// BARCODES DATA SETS ////////////////
// Sets for EAN13, EAN8, UPC-A, UPC-E
// all numbers are width. SetA and SetB : First number is Space, Second is Bar, Third is Space, Fourth is Bar. The SetC,wich is same as SetA but making Spaces instead of Bars and Bars instead of Spaces
var numberSetEAN13SetA = new Array(new Array(-3,2,-1,1),new Array(-2,2,-2,1),new Array(-2,1,-2,2),new Array(-1,4,-1,1),new Array(-1,1,-3,2),new Array(-1,2,-3,1),new Array(-1,1,-1,4),new Array(-1,3,-1,2),new Array(-1,2,-1,3),new Array(-3,1,-1,2));
var numberSetEAN13SetB = new Array(new Array(-1,1,-2,3),new Array(-1,2,-2,2),new Array(-2,2,-1,2),new Array(-1,1,-4,1),new Array(-2,3,-1,1),new Array(-1,3,-2,1),new Array(-4,1,-1,1),new Array(-2,1,-3,1),new Array(-3,1,-2,1),new Array(-2,1,-1,3));
var numberSetEAN13SetC = new Array(new Array(3,-2,1,-1),new Array(2,-2,2,-1),new Array(2,-1,2,-2),new Array(1,-4,1,-1),new Array(1,-1,3,-2),new Array(1,-2,3,-1),new Array(1,-1,1,-4),new Array(1,-3,1,-2),new Array(1,-2,1,-3),new Array(3,-1,1,-2));
var numberSetEAN13 = new Array(numberSetEAN13SetA, numberSetEAN13SetB, numberSetEAN13SetC);

// EAN13 : the below Array index is the check digit to encode
// each index contains 6 values to specify in which Set (A or B) each of the 6 first digits should be encoded with
// the value 0 or 1 inidcates Set A or B
var firstDigitEAN13Set = new Array(new Array(0,0,0,0,0,0),new Array(0,0,1,0,1,1),new Array(0,0,1,1,0,1),new Array(0,0,1,1,1,0),new Array(0,1,0,0,1,1),new Array(0,1,1,0,0,1),new Array(0,1,1,1,0,0),new Array(0,1,0,1,0,1),new Array(0,1,0,1,1,0),new Array(0,1,1,0,1,0));
// the 7 to 12 digits will be encoded using the SetC

//EAN13 XsizeRef is the base unit nominal width, also same for EAN8, UPC-A, UPC-E
var EAN13XsizeRef = 0.33; // size in millimeters
var EAN13BarcodeWidthRef = 31.35; // size in millimeters
var EAN13FullWidthRef = 37.29; // size in millimeters

// EAN13 structure
// LeftQuietZone, NormalGuardBarPattern, 6 digits from sets A or B, CenterGuardBarPattern, 6 digits from set C, NormalGuardBarPattern, RightQuietZone
// in these arrays, all numbers are the width (in base unit nominal width), NEGATIVE numbers are Spaces and POSITIVE numbers are Bars
var EAN13LeftQuietZone = new Array();
EAN13LeftQuietZone.push(-3.63/EAN13XsizeRef); // normal space of QuietZone = 3.63mm
var EAN13NormalGuardBarPattern = new Array(1,-1,1);
var EAN13CenterGuardBarPattern = new Array(-1,1,-1,1,-1);
var EAN13RightQuietZone = new Array();
EAN13RightQuietZone.push(-2.31/EAN13XsizeRef); // normal space of QuietZone = 2.31mm

var EAN13BarsHeight = 22.85; // height in millimeters
var EAN13GuardBarPatternBarsHeight = 24.5; // height in millimeters


// ITF14 XsizeRef is the narrow unit nominal width
var ITF14XsizeRef = 1.016; // size in millimeters
var ITF14NarrowToWideRatio = 2.5;
var ITF14BearerBarWidth = 4.83;
var ITF14FullWidthRef = ITF14XsizeRef * (7 * (4 * ITF14NarrowToWideRatio + 6) + ITF14NarrowToWideRatio + 6) + 2 * (10 * ITF14XsizeRef + ITF14BearerBarWidth); // size in millimeters
//$.writeln('ITF14FullWidthRef : '+ITF14FullWidthRef);


// ITF14 structure
// QuietZone, StartPattern, digits coded by pair, StopPattern, QuietZone
// QuietZone width = 10 *  ITF14XsizeRef

// ITF14 element widths Sets
//  each array is the digit to encode into 5 elements, the data 1 or ITF14NarrowToWideRatio is the width multiplier for the element
var ITF14digit0 = new Array(1,1,ITF14NarrowToWideRatio,ITF14NarrowToWideRatio,1);
var ITF14digit1 = new Array(ITF14NarrowToWideRatio,1,1,1,ITF14NarrowToWideRatio);
var ITF14digit2 = new Array(1,ITF14NarrowToWideRatio,1,1,ITF14NarrowToWideRatio);
var ITF14digit3 = new Array(ITF14NarrowToWideRatio,ITF14NarrowToWideRatio,1,1,1);
var ITF14digit4 = new Array(1,1,ITF14NarrowToWideRatio,1,ITF14NarrowToWideRatio);
var ITF14digit5 = new Array(ITF14NarrowToWideRatio,1,ITF14NarrowToWideRatio,1,1);
var ITF14digit6 = new Array(1,ITF14NarrowToWideRatio,ITF14NarrowToWideRatio,1,1);
var ITF14digit7 = new Array(1,1,1,ITF14NarrowToWideRatio,ITF14NarrowToWideRatio);
var ITF14digit8 = new Array(ITF14NarrowToWideRatio,1,1,ITF14NarrowToWideRatio,1);
var ITF14digit9 = new Array(1,ITF14NarrowToWideRatio,1,ITF14NarrowToWideRatio,1);
//  in these array, the index is the digit to encode
var numberSetITF14 = new Array(ITF14digit0,ITF14digit1,ITF14digit2,ITF14digit3,ITF14digit4,ITF14digit5,ITF14digit6,ITF14digit7,ITF14digit8,ITF14digit9);



///////////////////////////////////////////////////////////////

getDocumentScale();

getBarcodeObjectList();

checkBarcodeEditNeeded();

if (barcodeToEditList.length > 0){
    launchEdit();
}

if (correctedErrorList.length > 0){
    createReport();
}

function getDocumentScale(){
    // the sizes and coordinates in extendscript language are not stored in the document unit mode
    // as barcode standard are setted using millimeters, we need to convert all sizes and coordinates to millimeters for all the calculations
    
    var docUnit = new String(docObj.rulerUnits);
    docUnit = docUnit.split('.')[1];
    for (var i=0;i<unitTable[0].length && unitRatio == null;i++){
        if (unitTable[0][i] == docUnit){ 
            if (i > 1 && unitTable[1][i] != null){
                unitRatio = unitTable[1][i] * svgRatio;
            } else {
                $.writeln('ERROR - can\'t process document unit : '+docObj.rulerUnits); 
            }
        }
    }
    //$.writeln('unitRatio : '+unitRatio);
}    

function getBarcodeObjectList(){
    // check the entire document for existing supported "CoriusBarcode" groups
    var myGroup;
    var isBarcodeGroup;
    
    for (var i=0; i < docObj.groupItems.length; i++){
        isBarcodeGroup = false;
        myGroup = docObj.groupItems[i];
        if (myGroup.name != null && myGroup.name != ''){
            isBarcodeGroup = checkGroupNameIsCoriusBarcode(myGroup.name);
            
            if (isBarcodeGroup){
                barcodeObjectList.push(myGroup);
            }
        }
    }
}

function checkGroupNameIsCoriusBarcode(myName){
    // check each group name to see if contained in the supported "CoriusBarcode"  names list
    var result = false;
    var compareStr;
    
    for (var i=0; i<CoriusBarcodeNames.length; i++){
        compareStr = CoriusBarcodeNames[i];
        //$.writeln('CoriusBarcodeNames[j] : '+compareStr);
        if (myName == compareStr) {
            result = true;
            //$.writeln('CoriusBarcode FOUND : '+result);
        }
    }
    
    return result;
}

function checkBarcodeEditNeeded(){
    // check if the barcode needs to be redrawn from the code text in the settings
    var myGroup;
    var toEdit = false;
    
    for (var i=0; i<barcodeObjectList.length; i++){
        //$.writeln('iteration : '+i);
        myGroup = barcodeObjectList[i];
        toEdit = checkEditedText(myGroup);
        
        if (toEdit){
            barcodeToEditList.push(myGroup);
        }
    }
    //$.writeln('number of barcodes to edit found : '+barcodeToEditList.length); 
}

function checkEditedText(myGroup){
    var result = false;
    var currentCode = '';
    var myTextfield;
    var settingsCodeTFName = 'CODE_' + myGroup.name.substring(13,myGroup.name.length);
    var mySettingsGroup = myGroup.groupItems.getByName('Settings');
    var myBarcodeGroup = myGroup.groupItems.getByName('Barcode');
    //$.writeln('settingsCodeTFName : '+settingsCodeTFName);
    var wantedCode = mySettingsGroup.textFrames.getByName(settingsCodeTFName).contents;
    var currentCodeGroup = myBarcodeGroup.groupItems.getByName('HumanReadableCode');  
    
    for (var i=0; i < currentCodeGroup.textFrames.length; i++){
        myTextfield = currentCodeGroup.textFrames[i];
        currentCode += myTextfield.contents;
        //$.writeln('currentCode : adding '+myTextfield.contents);
    }
    
    if (currentCode != wantedCode){
        result = true;
    }
    
    return result;
}

function launchEdit(){
    //$.writeln('--- ENTERING EDIT STEPS --- ');
    var myItem;
    var settingsCodeTFName;
    var myBarcodeName;
    var myWantedCode;
    var myWantedCodeNoChecksum;
    var myCorrectedCode;
    var i;
    var myErrorArr;
    
    for (i=0;i<barcodeToEditList.length; i++){
        myErrorArr = new Array();
        myItem = barcodeToEditList[i];
        myBarcodeName = myItem.name;
        settingsCodeTFName = 'CODE_' + myBarcodeName.substring(13,myBarcodeName.length);
        myWantedCode = myItem.groupItems.getByName('Settings').textFrames.getByName(settingsCodeTFName).contents;
        
        // select the proper checking stream according to the type of barcode
        switch(myBarcodeName) {
          case 'CoriusBarcodeEAN13':
            myWantedCodeNoChecksum = myWantedCode.substr(0,12);
            myCorrectedCode = myWantedCodeNoChecksum + getChecksum(myWantedCodeNoChecksum);
            break;
          case 'CoriusBarcodeITF14':
            myWantedCodeNoChecksum = myWantedCode.substr(0,13);
            myCorrectedCode = myWantedCodeNoChecksum + getChecksum(myWantedCodeNoChecksum);
            break;
          default:
            $.writeln('no supported barcode format');
        } 
    
        if (myCorrectedCode != myWantedCode){
            // the barcode text in the settings isn't a proper code, the checksum digit doesn't match
            myErrorArr[0] = myWantedCode;     
            myErrorArr[1] = (myWantedCodeNoChecksum.length == myWantedCode.length)? '' : myWantedCode.substr( myWantedCode.length - 1, 1);
            myErrorArr[2] = myCorrectedCode.substr(myCorrectedCode.length - 1, 1);
            myErrorArr[3] = myCorrectedCode;
            myItem.groupItems.getByName('Settings').textFrames.getByName(settingsCodeTFName).contents = myCorrectedCode;
            
            correctedErrorList.push(myErrorArr); // this need to be implemented into a "result" .csv file to list all the errors and corrections.
        }
        
        // select the proper edition stream according to the type of barcode
        switch(myBarcodeName) {
          case 'CoriusBarcodeEAN13':
            drawEAN13(myItem);
            writeHumanReadableCode(myItem);
            myItem.groupItems.getByName('Settings').hidden = true;
            break;
          case 'CoriusBarcodeITF14':
            drawITF14(myItem);
            writeHumanReadableCode(myItem);
            myItem.groupItems.getByName('Settings').hidden = true;
            break;
          default:
            $.writeln('no supported barcode format');
        }
    }
}

function getChecksum(myCode){
    // this checksum algorythm is OK for : GTIN-8, GTIN-12, GTIN-13, GTIN-14, 17 digits, 18 digits
    var mySum = 0;
    var i;
    var multiplier = 1;
    var digit;
    var checksum;
    
    for (i=myCode.length -1; i >= 0; i--){
        multiplier = (multiplier == 1)? 3 : 1;
        digit = 1 * myCode.charAt(i);
        mySum += digit * multiplier;
        //$.writeln('multiplier / digit / mySum : '+multiplier+' / '+digit+' / '+mySum); 
    }
    if (mySum % 10 != 0){
        checksum = 10 - (mySum % 10);
        //$.writeln('--- CHECKED ---');
        //$.writeln('modulo10 / checksum : '+(mySum % 10)+' / '+checksum);
    } else {
        checksum = 0;
        //$.writeln('--- CHECKED ---');
        //$.writeln('already 10 multiple / checksum : '+(mySum % 10)+' / '+checksum);
    }
    
    return checksum;
}

function drawEAN13(myItem){
    //numberSetEAN13
    var myCodeTxt = myItem.groupItems.getByName('Settings').textFrames.getByName('CODE_EAN13').contents;
    //var myDigit;
    var myFirstDigit = 1 * myCodeTxt.substr(0,1);
    var i;
    var myFillColor = myItem.groupItems.getByName('Settings').textFrames.getByName('CODE_EAN13').textRanges[0].characterAttributes.fillColor;
    var myArrowSetting = myItem.groupItems.getByName('Settings').textFrames.getByName('ArrowYesNo').contents;
    var myBackgroundSetting = myItem.groupItems.getByName('Settings').textFrames.getByName('BackgroundYesNo').contents;
    var myBarsGrp = myItem.groupItems.getByName('Barcode').groupItems.getByName('Bars');
    
    var leftHalfSetsChoicesArray = firstDigitEAN13Set[myFirstDigit];    
    var SetsChoicesIndexes = new Array();
    var AllDigitsBarsSpaceWidth;
    
    var LeftGuardPlacer = myBarsGrp.pathItems.getByName('PlacerLeftGuardBars');
    var CenterGuardPlacer = myBarsGrp.pathItems.getByName('PlacerCenterGuardBars');
    var RightGuardPlacer = myBarsGrp.pathItems.getByName('PlacerRightGuardBars');
    var LeftHalfPlacer = myBarsGrp.pathItems.getByName('PlacerLeft');
    var RightHalfPlacer = myBarsGrp.pathItems.getByName('PlacerRight');
    
    var LeftDigitsBarsSpaceWidth;
    var RightDigitsBarsSpaceWidth;
    
    var myBackground = myItem.pathItems.getByName('OpaqueBackground');
    var myArrow = myItem.pathItems.getByName('Arrow');
    
    // setting the colors according to settings text colors
    myBackground.fillColor = myItem.groupItems.getByName('Settings').textFrames.getByName('BackgroundYesNo').textRanges[0].characterAttributes.fillColor;
    myArrow.fillColor = myFillColor;
    
    for (i=0; i<12; i++){
        if(i < 6){
            SetsChoicesIndexes.push(leftHalfSetsChoicesArray[i]);
        } else {
            SetsChoicesIndexes.push(2);
        }
    }

    // turning ON/OFF visibility according to settings
    myBackground.hidden = (myBackgroundSetting.toLowerCase() == 'no')? true : false;
    myArrow.hidden = (myArrowSetting.toLowerCase() == 'no')? true : false;
    
    // Delete all existing bars from the Barcode group for a clean start
    cleanBars(myBarsGrp);
    
    AllDigitsBarsSpaceWidth = getAllWidthsEAN13(myCodeTxt,SetsChoicesIndexes,numberSetEAN13);
    LeftDigitsBarsSpaceWidth = AllDigitsBarsSpaceWidth.slice(0,6);
    RightDigitsBarsSpaceWidth = AllDigitsBarsSpaceWidth.slice(6,12);
    
    // draw Guard Bars Pattern    
    var myAngle = calculAngle(LeftGuardPlacer);
    var myNewGrp;    
    docObj.selection = null;
    
    drawBars(myBarsGrp,LeftGuardPlacer,myFillColor,EAN13NormalGuardBarPattern,EAN13FullWidthRef,0,myAngle);    
    drawBars(myBarsGrp,RightGuardPlacer,myFillColor,EAN13NormalGuardBarPattern,EAN13FullWidthRef,0,myAngle);
    drawBars(myBarsGrp,CenterGuardPlacer,myFillColor,EAN13CenterGuardBarPattern,EAN13FullWidthRef,0,myAngle);
    
    // draw Code Bars
    prepDrawBarsEAN13(myBarsGrp,LeftHalfPlacer,myFillColor,LeftDigitsBarsSpaceWidth,EAN13FullWidthRef,myAngle);
    prepDrawBarsEAN13(myBarsGrp,RightHalfPlacer,myFillColor,RightDigitsBarsSpaceWidth,EAN13FullWidthRef,myAngle);
    
    docObj.selection = null;
}

function drawITF14(myItem){
    var myCodeTxt = myItem.groupItems.getByName('Settings').textFrames.getByName('CODE_ITF14').contents;
    var i;
    var j;
    var myFillColor = myItem.groupItems.getByName('Settings').textFrames.getByName('CODE_ITF14').textRanges[0].characterAttributes.fillColor;
    var myBackgroundSetting = myItem.groupItems.getByName('Settings').textFrames.getByName('BackgroundYesNo').contents;
    var myFrameSetting = myItem.groupItems.getByName('Settings').textFrames.getByName('FrameYesNo').contents;
    var myBarsGrp = myItem.groupItems.getByName('Barcode').groupItems.getByName('Bars');
    var myPlacer = myBarsGrp.pathItems.getByName('Placer');
    var myWidthsArr = new Array();
    var myDigit1;
    var myDigit2;
    var myWlist1;
    var myWlist2;
    var myW1;
    var myW2;
    var myAngle;
    
    var myBackground = myItem.pathItems.getByName('OpaqueBackground');
    var myFrame = myItem.compoundPathItems.getByName('BearerBars');
    var myPermBars = myItem.groupItems.getByName('Barcode').groupItems.getByName('PermanentBars');
    
    // setting the colors according to settings text colors
    myBackground.fillColor = myItem.groupItems.getByName('Settings').textFrames.getByName('BackgroundYesNo').textRanges[0].characterAttributes.fillColor;
    myFrame.pathItems[0].fillColor = myFillColor;
    for (i=0;i<myPermBars.pathItems.length;i++){
        myPermBars.pathItems[i].fillColor = myFillColor;
    }
    // turning ON/OFF visibility according to settings
    myBackground.hidden = (myBackgroundSetting.toLowerCase() == 'no')? true : false;
    myFrame.hidden = (myFrameSetting.toLowerCase() == 'no')? true : false;
    
    for (i=0;i<myCodeTxt.length-1;i+=2){
        myDigit1 = 1 * myCodeTxt[i];
        myDigit2 = 1 * myCodeTxt[i+1];
        myWlist1 = numberSetITF14[myDigit1];
        myWlist2 = numberSetITF14[myDigit2];
        for (j=0;j<myWlist1.length;j++){
            myW1 = myWlist1[j];
            myW2 = myWlist2[j];
            myWidthsArr.push(myW1);
            myWidthsArr.push(-myW2);
        }
    }

    myAngle = calculAngle(myPlacer);
    
    // Delete all existing bars from the Barcode group for a clean start
    cleanBars(myBarsGrp);
    
    var myNewGrp;    
    docObj.selection = null;
    
    drawBars(myBarsGrp,myPlacer, myFillColor, myWidthsArr,ITF14FullWidthRef,0,myAngle);
    docObj.selection = null;
}

function cleanBars(myBarsGrp){
    // delete all existing bars
    var i;
    var j = 0;
    var myObj;
    
    for (i=myBarsGrp.pathItems.length; i>0; i--){
        myObj = myBarsGrp.pathItems[i-1];
        if(myObj.name == null || myObj.name == ''){
            myObj.remove();
        }
    }
}

function getAllWidthsEAN13(myCodeTxt,mySetsChoicesIndexes,myNumberSets){
    var result = new Array();
    var myDigit;
    var i;
    var mySetIndex;
    var mySet;
    var myWidth;
    
    for (i=1; i < myCodeTxt.length; i++){
        myDigit = 1 * myCodeTxt.substr(i,1);
        mySetIndex = mySetsChoicesIndexes[i-1];
        mySet = myNumberSets[mySetIndex];
        myWidth = mySet[myDigit];
        
        result.push(myWidth);
    }
    
    return result;
}

function prepDrawBarsEAN13(myBarsGrp,myPlaceHolder, myFillColor, myWidthsArr,myNormalWidth,myAngle){
    var i;
    var myScale = (calculFlatWidth(myBarsGrp.parent.parent.pathItems.getByName('OpaqueBackground'),myAngle) * unitRatio) / myNormalWidth;
    var myDigitWidthsArr;
    var decalage = 0;
    
    for (i=0; i<myWidthsArr.length; i++){
        myDigitWidthsArr = myWidthsArr[i];
        //$.writeln('myWidthsArr['+i+'] : '+myWidthsArr[i]);
        drawBars(myBarsGrp,myPlaceHolder,myFillColor,myDigitWidthsArr,EAN13FullWidthRef,decalage,myAngle);
        decalage += 7  * EAN13XsizeRef * myScale / unitRatio;
    }
}

function drawBars(myBarsGrp,myPlaceHolder, myFillColor, myWidthsArr,myNormalWidth,decalage,myAngle){
    //EAN13XsizeRef
    var i;
    var myNormalizedW;
    var myFlatW;
    var myRealW;
    var myScale = (calculFlatWidth(myBarsGrp.parent.parent.pathItems.getByName('OpaqueBackground'),myAngle) * unitRatio) / myNormalWidth;
    var myRefLowPoint = myPlaceHolder.pathPoints[1];
    var myRefHighPoint = myPlaceHolder.pathPoints[2];
    
    myScale = Math.round(myScale * 10000)/10000;
    
    var myBarcodeName = myBarsGrp.parent.parent.name;
    
    myBarsGrp.parent.locked = false;
    myBarsGrp.locked = false;
    
    var decalX = decalage * Math.cos(Math.PI * myAngle / 180);
    var decalY = decalage * Math.sin(Math.PI * myAngle / 180);
    var myBar;
    var bottomLeft;
    var bottomRight;
    var topRight;
    var topLeft;
    var SIZEREF;
    
    switch(myBarcodeName) {
        case 'CoriusBarcodeEAN13':
            SIZEREF = EAN13XsizeRef;
            break;
        case 'CoriusBarcodeITF14':
            SIZEREF = ITF14XsizeRef;
            break;
        default:
            $.writeln('no supported barcode format');
    }
    //$.writeln('SIZEREF : '+SIZEREF);
    
    for (i=0;i<myWidthsArr.length;i++){
        myBar = null;
        myNormalizedW = 1 * myWidthsArr[i];
        myRealW = Math.abs(myNormalizedW) * SIZEREF * myScale / unitRatio;
        myFlatW = myRealW * Math.cos(Math.PI * myAngle / 180);
        if (myNormalizedW > 0){            
            myBar = myBarsGrp.pathItems.add();
            
            bottomLeft = myBar.pathPoints.add();
            bottomLeft.anchor = [myRefLowPoint.anchor[0] + decalX, myRefLowPoint.anchor[1] + decalY];
            bottomLeft.leftDirection = bottomLeft.anchor;
            bottomLeft.rightDirection = bottomLeft.anchor;
            bottomLeft.pointType = PointType.CORNER;
            
            topLeft = myBar.pathPoints.add();
            topLeft.anchor = [myRefHighPoint.anchor[0] + decalX, myRefHighPoint.anchor[1] + decalY];
            topLeft.leftDirection = topLeft.anchor;
            topLeft.rightDirection = topLeft.anchor;
            topLeft.pointType = PointType.CORNER; 
            
            topRight = myBar.pathPoints.add();
            topRight.anchor = [myRefHighPoint.anchor[0] + decalX + myFlatW, myRefHighPoint.anchor[1] + decalY + myRealW * Math.sin(Math.PI * myAngle / 180)];
            topRight.leftDirection = topRight.anchor;
            topRight.rightDirection = topRight.anchor;
            topRight.pointType = PointType.CORNER; 
            
            bottomRight = myBar.pathPoints.add();
            bottomRight.anchor = [myRefLowPoint.anchor[0] + decalX + myFlatW, myRefLowPoint.anchor[1] + decalY + myRealW * Math.sin(Math.PI * myAngle / 180)];
            bottomRight.leftDirection = bottomRight.anchor;
            bottomRight.rightDirection = bottomRight.anchor;
            bottomRight.pointType = PointType.CORNER;
            
            myBar.filled = true;
            myBar.closed = true;
            myBar.fillColor = myFillColor;
            myBar.stroked = false;
            
            myBar.selected = true;
        } 
        decalX += myFlatW;
        decalY += myRealW * Math.sin(Math.PI * myAngle / 180);
    }
    
}

function calculFlatWidth(myPlaceHolder,myAngle){
    // to obtain the actual width of the barcode, regardless of its rotation
    var result;
    var myX1 = myPlaceHolder.pathPoints[1].anchor[0];
    var myX2 = myPlaceHolder.pathPoints[2].anchor[0];
    var gapX = myX2 - myX1;
    
    result = gapX / Math.cos(Math.PI * myAngle / 180);
    result = Math.round(result * 10000)/10000;

    return result;
}

function calculAngle(myPlaceHolder){
    // to obtain the rotation angle of the barcode
    var p0 = myPlaceHolder.pathPoints[0];
    var p1 = myPlaceHolder.pathPoints[3];
    var result;
    
    var deltaX = p0.anchor[0] - p1.anchor[0];
    var deltaY = p0.anchor[1] - p1.anchor[1];
    
    result = -180 * Math.atan(deltaX / deltaY) / Math.PI;

    return result;
}

function writeHumanReadableCode(myCoriusBarcode){
    var settingsCodeTFName = 'CODE_' + myCoriusBarcode.name.substring(13,myCoriusBarcode.name.length);
    var myGrp = myCoriusBarcode.groupItems.getByName('Barcode').groupItems.getByName('HumanReadableCode');
    var myCodeTxt = myCoriusBarcode.groupItems.getByName('Settings').textFrames.getByName(settingsCodeTFName).contents;
    var myFillColor = myCoriusBarcode.groupItems.getByName('Settings').textFrames.getByName(settingsCodeTFName).textRanges[0].characterAttributes.fillColor;
    var myTxtField;
    var i; 
    var startId = 0;
    var endId;
    var current_paragraphs;
    
    for (i=0;i<myGrp.textFrames.length; i++){
        myTxtField = myGrp.textFrames[i];
        endId = startId + myTxtField.contents.length;
        //$.writeln('myTxtField.name : '+myTxtField.name);
        //$.writeln('myTxtField.zOrderPosition : '+myTxtField.zOrderPosition);
        myTxtField.contents = myCodeTxt.substring(startId,endId);
        
        current_paragraphs = myTxtField.textRange.paragraphs;
        current_paragraphs[0].fillColor = myFillColor;
        //myTxtField.characters.characterAttributes.fillColor = myFillColor;
        startId = endId;
    }        
}

function createReport(){
    var myFileName = docObj.name.substring(0,docObj.name.lastIndexOf ('.ai'))+'_'+'CoriusBarcode_REPORT.csv';
    var myFile;
    var myFolderName = docFolder;
    var myFolder = new Folder(myFolderName);
    var i;
    
    for (i=0; i<correctedErrorList.length; i++){
        CSVReportTxt += correctedErrorList[i].join(';');
        CSVReportTxt += '\n';
    }
    $.writeln('CSVReportTxt : '+CSVReportTxt);
    
    // uncomment below line if the report needs to be in a subfolder
    //myFolder.create();
        
    myFile = new File(myFolderName+'\\'+myFileName);
    myFile.encoding = "BINARY";
    myFile.open('w');
    myFile.write(CSVReportTxt);
    myFile.close();    
    
    alert('Checksum error(s) found ! \nCheck the report here : '+myFile);
}