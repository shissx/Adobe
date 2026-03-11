app.executeMenuCommand('doc-color-cmyk');

var blackColor = getCmykColor(0, 0, 0, 100);
var cyanColor = getCmykColor(100, 0, 0, 0);
var cyanColor10 = getCmykColor(10, 0, 0, 0);
var cyanColor15 = getCmykColor(15, 0, 0, 0);
var cyanColor20 = getCmykColor(20, 0, 0, 0);
var cyanColor25 = getCmykColor(25, 0, 0, 0);
var cyanColor30 = getCmykColor(30, 0, 0, 0);

var myDoc = app.activeDocument;

//处理文字-颜色
for (var i = 0; i < myDoc.textFrames.length; i++) {
    var myTF = myDoc.textFrames[i];
    //alert (myTF.contents)
    for ( j = 0; j < myTF.words.length; j++) {
        word = myTF.words[j];
        if ( word.fillColor.cyan > 60 && word.fillColor.black < 5 ) {
            word.filled = true;
            word.fillColor = cyanColor;
        } else {
            myTF.textRange.characterAttributes.fillColor = blackColor;//全部替换颜色
        }
    }
}

//转曲2
/*
#target Illustrator
#targetengine main
app.activeDocument.selection = null;  
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
var myDoc = app.activeDocument;
//
// ALL TEXT create outlines (ie. VECTORIZED all text)
//var textToVectorized = myDoc.textFrames;
for ( var tf = textToVectorized.length-1; tf >= 0; tf-- ) {
var text = textToVectorized[tf];
text.createOutline();
}
*/

//处理图形
for (var i = 0; i < myDoc.pathItems.length; i++) {
    var pathItemfillColor = myDoc.pathItems[i].fillColor;
    var pathItemstrokeColor = myDoc.pathItems[i].strokeColor;
    //填充色
    if (pathItemfillColor == "[CMYKColor]") {
        if (pathItemfillColor.cyan == 0 && pathItemfillColor.magenta == 0 && pathItemfillColor.yellow == 0 &&
            pathItemfillColor.black == 0) {
            //alert("CMYK→无填充")
        } else {
            //alert("CMYK→有填充")
            myDoc.pathItems[i].fillColor = cyanColor30;
        }
    } else if (pathItemfillColor == "[GrayColor]") { //灰度即为无填充
        if (pathItemfillColor.gray == 0) {
            //alert("灰度→无填充")
        } else {
            //alert("灰度→有填充")
        }
    }
    //描边色
    if (pathItemstrokeColor == "[CMYKColor]") {
        if (pathItemfillColor.cyan == 0 && pathItemfillColor.magenta == 0 && pathItemfillColor.yellow == 0 &&
            pathItemfillColor.black == 0) {
            //alert("CMYK→无填充")
        } else {
            //alert("CMYK→有填充")
            myDoc.pathItems[i].strokeColor = cyanColor;
        }
    } else if (pathItemstrokeColor == "[GrayColor]") { //灰度即为无填充
        if (pathItemfillColor.gray == 0) {
            //alert("灰度→无填充")
        } else {
            //alert("灰度→有填充")
        }
    }
}

//处理文字-转曲
while (myDoc.textFrames.length != 0) {
    myDoc.textFrames[0].createOutline(); 
}

//另存为PDF
var fpath = myDoc.path; 
savePDF( fpath )
function savePDF( file ) {
	var saveOpts = new PDFSaveOptions();
	saveOpts.compatibility = PDFCompatibility.ACROBAT4;//兼容PDF1.3
    saveOpts.preserveEditability = true;//保留AI编辑功能	
    saveOpts.generateThumbnails = true;//嵌入页面缩略图
    myDoc.saveAs( file, saveOpts );
}
alert ("转双色完成！","正仁集团")

function getCmykColor(c, m, y, k) {
    newCMYKColor = new CMYKColor();
    newCMYKColor.black = k;
    newCMYKColor.cyan = c;
    newCMYKColor.magenta = m;
    newCMYKColor.yellow = y;
    return newCMYKColor;
}