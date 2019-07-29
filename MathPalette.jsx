/*
Palette for math illustrations in Illustrator.
Requires pdflatex and ghostscript > 9.15.
These must be in the path.

Inspiration from Gilles Castel (https://castel.dev/post/lecture-notes-2/)
Code for adding items from Harmut Lemmel on the (now-defunct) latex-community.org site.
https://web.archive.org/web/20170113205324/http://latex-community.org/know-how/latexs-friends/61-latexs-friends-others/381-combining-latex-and-illustrator
*/

#target Illustrator
#targetengine main

var pdflatexexe="pdflatex.exe";
var temppath=Folder.temp.fsName;

var palette = new Window('palette', "Math");
palette.preferredSize.width=200;
palette.text1 = palette.add('edittext', undefined, "$F = m a$");
palette.text1.preferredSize.width = 190;
palette.btn1 = palette.add('button', undefined, 'Add TeX', {name: 'add'});
palette.fillGroup = palette.add("group");
palette.fillGroup.orientation = "row";
palette.fillGroup.fillwhitebtn = palette.fillGroup.add('button', undefined, 'Fill White');
palette.fillGroup.fillgraybtn = palette.fillGroup.add('button', undefined, 'Fill Gray');
palette.fillGroup2 = palette.add("group");
palette.fillGroup2.orientation = "row";
palette.fillGroup2.fillblackbtn = palette.fillGroup2.add('button', undefined, 'Fill Black');
palette.fillGroup2.fillnonebtn = palette.fillGroup2.add('button', undefined, 'Fill None');
palette.strokeGroup = palette.add("group");
palette.strokeGroup.orientation = "row";
palette.strokeGroup.strokeSolidBtn = palette.strokeGroup.add('button', undefined, 'Stroke Solid');
palette.strokeGroup.strokeDotBtn = palette.strokeGroup.add('button', undefined, 'Stroke Dotted');
palette.strokeGroup2 = palette.add("group");
palette.strokeGroup2.orientation = "row";
palette.strokeGroup2.strokeDashBtn = palette.strokeGroup2.add('button', undefined, 'Stroke Dashed');
palette.strokeGroup2.strokeNoneBtn = palette.strokeGroup2.add('button', undefined, 'Stroke None');

palette.btn1.onClick = function(){
    var loadCode = ' \
        // import pdf file into the current document \
        var temppath=Folder.temp.fsName; \
        var pdffile=File(temppath+"\\latex2illustrator.pdf"); \
        var grp=app.activeDocument.activeLayer.groupItems.createFromFile(pdffile); \
        // The imported objects are grouped twice. Now move the subgroup \
        // items to the main group and skip the last item which is the page frame \
        for( var i=grp.pageItems[0].pageItems.length; --i>=0; ) \
        { \
            grp.pageItems[0].pageItems[i].move(grp,ElementPlacement.PLACEATEND); \
        } \
        var last = grp.pageItems.length - 1; \
        if (last >= 0 && grp.pageItems[last].typename == "PathItem") \
        { \
            grp.pageItems[last].remove(); \
        } \
        // Move the imported objects to the center of the current view. \
        grp.translate(app.activeDocument.activeView.centerPoint[0]-grp.left, app.activeDocument.activeView.centerPoint[1]-grp.top); \
        ';
    var latexfile=new File(temppath+'\\latex2illustrator.tex');
    var latexcode = palette.text1.text;
    latexfile.open("w");
    latexfile.writeln("\\documentclass[12pt]{standalone}");
    // add or remove additional latex packages here
    latexfile.writeln("\\usepackage{amsmath}");
    latexfile.writeln("\\usepackage{amssymb}");
    latexfile.writeln("\\usepackage{bm}");        // bold math
    latexfile.writeln("\\begin{document}");
    latexfile.writeln("\\pagestyle{empty}"); // no page number
    latexfile.writeln(latexcode);
    latexfile.writeln("\\end{document}");
    latexfile.close();
    
    var pdffile=File(temppath+"\\latex2illustrator.pdf");
    if(pdffile.exists)
    {
        pdffile.remove();
    }
        
    // create a batch file calling latex
    var batchfile=new File(temppath+'\\latex2illustrator.bat');
    batchfile.open("w");
    batchfile.writeln(pdflatexexe+' -aux-directory="'+temppath+'" -include-directory="'+temppath+'" -output-directory="'+temppath+'" "'+temppath+'\\latex2illustrator.tex"');
    batchfile.writeln('gswin64c -o "'+temppath+'\\latex2illustrator2.pdf' +'" -dNoOutputFonts -sDEVICE=pdfwrite "'+temppath+'\\latex2illustrator.pdf"');
    batchfile.writeln('del "'+temppath+'\\latex2illustrator.pdf"');
    batchfile.writeln('ren "'+temppath+'\\latex2illustrator2.pdf" "latex2illustrator.pdf"');
    batchfile.writeln('del "'+temppath+'\\latex2illustrator.bat"');
    batchfile.close();
    
    batchfile.execute();
    var waited = 0;
    // wait until the batch file has removed itself
    while(batchfile.exists && waited < 5000)
    {
        $.sleep(100);
        waited += 100;
    }    
    if(waited >= 5000)    
        alert('batch file timed out');            
    
    var pdffile=File(temppath+"\\latex2illustrator.pdf");
    if(pdffile.exists)
    {
        sendToIllustrator(loadCode);
    }
    else
    {        
        alert("File "+temppath+"\\"+pdffile.name+" could not be created. LaTeX error?");
    }
}

palette.fillGroup.fillwhitebtn.onClick = function(){
    var loadcode = fillSelectedWithColor(255, 255, 255);
    sendToIllustrator(loadcode);
}

palette.fillGroup.fillgraybtn.onClick = function(){
    var loadcode = fillSelectedWithColor(218, 218, 218);
    sendToIllustrator(loadcode);
}

palette.fillGroup2.fillblackbtn.onClick = function(){
    var loadcode = fillSelectedWithColor(0, 0, 0);
    sendToIllustrator(loadcode);
}

palette.fillGroup2.fillnonebtn.onClick = function(){
    var loadcode = 'var selected = app.selection; \
    for(var i=0; i < selected.length; i++) \
    { \
        selected[i].filled = false; \
    }'
    sendToIllustrator(loadcode);
}

palette.strokeGroup.strokeSolidBtn.onClick = function(){
    var loadcode = 'var selected = app.selection;\
    var color = new RGBColor();\
    color.red = 0;\
    color.green = 0;\
    color.blue = 0;\
    for(var i=0; i < selected.length; i++)\
    {\
        if(!selected[i].stroked) selected[i].stroked = true;\
        selected[i].strokeDashes = [];\
    }';
    sendToIllustrator(loadcode);
}

palette.strokeGroup.strokeDotBtn.onClick = function(){
    var loadcode = 'var selected = app.selection;\
    for(var i=0; i < selected.length; i++)\
    {\
        if(!selected[i].stroked) selected[i].stroked = true;\
        selected[i].strokeDashes = [3, 6];\
    }';
    sendToIllustrator(loadcode);
}

palette.strokeGroup2.strokeDashBtn.onClick = function(){
    var loadcode = 'var selected = app.selection;\
    for(var i=0; i < selected.length; i++)\
    {\
        if(!selected[i].stroked) selected[i].stroked = true;\
        selected[i].strokeDashes = [8];\
    }';
    sendToIllustrator(loadcode);
}

palette.strokeGroup2.strokeNoneBtn.onClick = function(){
    var loadcode = 'var selected = app.selection;\
    for(var i=0; i < selected.length; i++)\
    {\
        if(selected[i].stroked) selected[i].stroked = false;\
    }';
    sendToIllustrator(loadcode);
}

function fillSelectedWithColor(red, green, blue)
{
    var s = 'var selected = app.selection; \
    var color = new RGBColor();';
    s += 'color.red = ' + red + ';';
    s += 'color.green = ' + green + ';';
    s += 'color.blue = ' + blue + ';';
    s += ' \
    function fillPaths(item, color) \
    { \
        if(item.typename == "PathItem") \
        {\
            item.fillColor = color; \
            return; \
        } \
        if(item.typename == "GroupItem" || item.typename == "CompoundPathItem") \
        { \
            if(item.groupItems) \
            { \
                for(var i=0; i < item.groupItems.length; i++) \
                    fillPaths(item.groupItems[i], color); \
            } \
            if(item.compoundPathItems) \
            { \
                for(var i=0; i < item.compoundPathItems.length; i++) \
                    fillPaths(item.compoundPathItems[i], color); \
            } \
            if(item.pathItems) \
            { \
                for(var i=0; i < item.pathItems.length; i++) \
                    fillPaths(item.pathItems[i], color); \
            } \
        } \
    } \
    for(var i=0; i < selected.length; i++) \
    { \
        fillPaths(selected[i], color); \
    }';
    return s;
}

function sendToIllustrator(loadcode)
{
    var bt = new BridgeTalk;  
    bt.target = "Illustrator";  
    bt.body = loadcode;
    bt.send();
}

palette.show();