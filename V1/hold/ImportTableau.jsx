/*
  Script Adobe Illustrator : Import CSV -> Tableau (V7 Compact & Silent Mode) + PREVIEW FIXÉ
*/
#targetengine "session"

var previewGroup = null; 

// --- FONCTIONS COULEURS ---
function makeCMYK(c, m, y, k) {
    var color = new CMYKColor();
    color.cyan = parseFloat(c); color.magenta = parseFloat(m);
    color.yellow = parseFloat(y); color.black = parseFloat(k);
    return color;
}

function cmykToScreenRGB(c, m, y, k) {
    var C = c/100; var M = m/100; var Y = y/100; var K = k/100;
    var r = 1 - Math.min(1, C*(1-K)+K);
    var g = 1 - Math.min(1, M*(1-K)+K);
    var b = 1 - Math.min(1, Y*(1-K)+K);
    return [r, g, b];
}

// --- NETTOYAGE PRÉVISUALISATION FIABLE ---
function cleanupPreview() {
    try {
        if (previewGroup && previewGroup.isValid) {
            try { previewGroup.remove(); } catch(e){}
        }
        previewGroup = null;

        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var items = doc.pageItems;

        for (var i = items.length - 1; i >= 0; i--) {
            try {
                if (items[i].name === "Tableau_PREVIEW_TEMP") {
                    items[i].remove();
                }
            } catch(e){}
        }
    } catch(e){}
}

// --- UI COMPACTE ---
function addCompactSlider(parent, label, initVal, callback) {
    var g = parent.add("group");
    g.orientation = "row"; g.alignChildren = ["left", "center"]; g.spacing = 2;
    g.margins = 0;

    var lbl = g.add("statictext", undefined, label); 
    lbl.preferredSize.width = 10;

    var sld = g.add("slider", undefined, parseInt(initVal), 0, 100); 
    sld.preferredSize.width = 60;

    var inp = g.add("edittext", undefined, initVal); 
    inp.preferredSize.width = 25;

    sld.onChanging = function() { inp.text = Math.round(this.value); callback(); };
    inp.onChange = function() {
        var v = parseFloat(this.text); if(isNaN(v)) v=0; if(v<0) v=0; if(v>100) v=100;
        this.text = v; sld.value = v; callback();
    };
    return { slider: sld, input: inp };
}

function createCompactColorBlock(parentTab, title, defC, defM, defY, defK) {
    var grpLine = parentTab.add("group");
    grpLine.orientation = "row"; grpLine.alignChildren = ["left", "center"]; 
    grpLine.spacing = 5;

    var lblTitle = grpLine.add("statictext", undefined, title);
    lblTitle.preferredSize.width = 50;

    var previewBox = grpLine.add("panel"); 
    previewBox.preferredSize = [20, 40];

    var grpSliders = grpLine.add("group"); 
    grpSliders.orientation = "column"; 
    grpSliders.spacing = 0; 
    grpSliders.margins = 0;

    var ui = {}; 
    var updatePreview = function() {
        if (!ui.c) return; 
        var rgb = cmykToScreenRGB(ui.c.input.text, ui.m.input.text, ui.y.input.text, ui.k.input.text);
        var g = previewBox.graphics;
        var brush = g.newBrush(g.BrushType.SOLID_COLOR, [rgb[0], rgb[1], rgb[2], 1]);
        g.backgroundColor = brush;
        previewBox.visible = false; previewBox.visible = true; 
    };

    ui.c = addCompactSlider(grpSliders, "C", defC, updatePreview);
    ui.m = addCompactSlider(grpSliders, "M", defM, updatePreview);
    ui.y = addCompactSlider(grpSliders, "J", defY, updatePreview);
    ui.k = addCompactSlider(grpSliders, "N", defK, updatePreview);
    updatePreview();
    return ui; 
}

// --- MAIN ---
function main() {

    // Anti "No Document"
    if (app.documents.length === 0) {
        alert("Ouvrez un document Illustrator avant d'utiliser ce script.");
        return;
    }

    var csvFile;
    if ($.os.indexOf("Windows") !== -1) csvFile = File.openDialog("Sélectionnez CSV", "*.csv;*.txt;*.*"); 
    else csvFile = File.openDialog("Sélectionnez CSV", function(f) { return true; }); 
    if (!csvFile) return;

    csvFile.open("r"); var content = csvFile.read(); csvFile.close();
    var lines = content.split(/\r\n|\r|\n/);
    if (lines.length < 1) return;

    var separator = (lines[0].indexOf(";") > -1) ? ";" : ",";
    var data = [];
    for (var i = 0; i < lines.length; i++) {
        if (lines[i] !== "") data.push(lines[i].split(separator));
    }
    if(data.length === 0) return;
    var headers = data[0]; 

    // --- FENETRE COMPACTE ---
    var win = new Window("dialog", "Tableau V7 (Compact)");

    win.orientation = "column"; win.spacing = 5; win.margins = 5;

    var grpMain = win.add("group"); 
    grpMain.orientation = "row"; grpMain.alignChildren = ["left", "top"]; grpMain.spacing = 10;

    // COLONNE GAUCHE
    var colLeft = grpMain.add("group"); colLeft.orientation = "column"; colLeft.spacing = 2;

    var pGeo = colLeft.add("panel", undefined, "Géométrie"); 
    pGeo.alignChildren = "left"; pGeo.margins = 5;

    var addSmallInput = function(p, lbl, val) {
        var g = p.add("group"); g.spacing = 2;
        g.add("statictext", undefined, lbl).preferredSize.width = 45;
        var i = g.add("edittext", undefined, val); i.preferredSize.width = 30;
        return i;
    };

    var inpHeadH = addSmallInput(pGeo, "H.Titre:", "40");
    var inpRowH  = addSmallInput(pGeo, "H.Ligne:", "20");
    var inpPad   = addSmallInput(pGeo, "Marge:", "15");

    // Typo
    var pTypo = colLeft.add("panel", undefined, "Typo"); 
    pTypo.alignChildren = "left"; pTypo.margins = 5;

    var gF = pTypo.add("group"); gF.spacing = 2;
    gF.add("statictext", undefined, "Font:");
    var inpFont = gF.add("edittext", undefined, "ArialMT"); 
    inpFont.preferredSize.width = 60;

    var inpSize = addSmallInput(pTypo, "Taille:", "10");
    var chkBold = pTypo.add("checkbox", undefined, "Gras (Titre)");

    // Colonnes
    var pCols = colLeft.add("panel", undefined, "Colonnes");
    pCols.margins = 5;

    var colScroll = pCols.add("group"); 
    colScroll.orientation = "column"; colScroll.alignChildren = "left"; colScroll.spacing = 0;

    var checkBoxes = [];
    for (var h = 0; h < headers.length; h++) {
        var cb = colScroll.add("checkbox", undefined, headers[h]); 
        cb.value = true; checkBoxes.push(cb);
    }

    // COLONNE DROITE
    var colRight = grpMain.add("group"); colRight.orientation = "column";

    var tPanel = colRight.add("tabbedpanel"); 
    tPanel.preferredSize = [220, 250];

    var tabHead = tPanel.add("tab", undefined, "Entête");
    tabHead.alignChildren = ["fill", "top"]; tabHead.margins = 5;
    var uiTxtHead = createCompactColorBlock(tabHead, "Texte", 0, 0, 0, 0);
    var uiBgHead = createCompactColorBlock(tabHead, "Fond", 0, 0, 0, 80);

    var tabCont = tPanel.add("tab", undefined, "Contenu");
    tabCont.alignChildren = ["fill", "top"]; tabCont.margins = 5;
    var uiTxtCont = createCompactColorBlock(tabCont, "Texte", 0, 0, 0, 100);
    var uiBgRow1 = createCompactColorBlock(tabCont, "Ligne 1", 0, 0, 0, 0);
    var uiBgRow2 = createCompactColorBlock(tabCont, "Ligne 2", 0, 0, 0, 10);
    tPanel.selection = 0;

    // Boutons
    var gBtn = win.add("group"); gBtn.spacing = 5;
    var btnPreview = gBtn.add("button", undefined, "Preview"); 
    btnPreview.preferredSize.width = 70;
    var btnOK = gBtn.add("button", undefined, "OK");
    btnOK.preferredSize.width = 50;
    var btnCancel = gBtn.add("button", undefined, "X");
    btnCancel.preferredSize.width = 30;

    // LOGIQUE
    var runDraw = function(isPreview) {
        if (app.documents.length === 0) {
            alert("Ouvrez un document avant d'utiliser ce script.");
            return;
        }

        var colors = {
            txtHead: [uiTxtHead.c.input.text, uiTxtHead.m.input.text, uiTxtHead.y.input.text, uiTxtHead.k.input.text],
            bgHead:  [uiBgHead.c.input.text, uiBgHead.m.input.text, uiBgHead.y.input.text, uiBgHead.k.input.text],
            txtCont: [uiTxtCont.c.input.text, uiTxtCont.m.input.text, uiTxtCont.y.input.text, uiTxtCont.k.input.text],
            bgRow1:  [uiBgRow1.c.input.text, uiBgRow1.m.input.text, uiBgRow1.y.input.text, uiBgRow1.k.input.text],
            bgRow2:  [uiBgRow2.c.input.text, uiBgRow2.m.input.text, uiBgRow2.y.input.text, uiBgRow2.k.input.text]
        };

        if (isPreview) cleanupPreview();

        dessinerTableau(data, checkBoxes, inpHeadH.text, inpRowH.text, inpPad.text, 
            inpFont.text, inpSize.text, chkBold.value, colors, isPreview);

        try { app.redraw(); } catch(e){}
    };

    btnPreview.onClick = function() {
        if (app.documents.length === 0) { alert("Aucun document ouvert."); return; }
        cleanupPreview();
        runDraw(true);
    };

    btnOK.onClick = function() {
        if (previewGroup && previewGroup.isValid) {
            previewGroup.name = "Tableau Final";
            previewGroup = null;
        } else {
            runDraw(false);
        }
        win.close();
    };

    btnCancel.onClick = function() { cleanupPreview(); win.close(); };
    win.show();
    app.activate();
    win.active = true;
}

// --- MOTEUR DE DESSIN ---
function dessinerTableau(data, checkBoxes, hHeadVal, hRowVal, padVal, fontName, fSizeVal, isBold, uiColors, isPreview) {

    if (app.documents.length === 0) {
        alert("Aucun document ouvert.");
        return;
    }

    var doc = app.activeDocument;

    try {
        var layerName = "Calque_Tableau";
        var layer;
        try { layer = doc.layers.getByName(layerName); } 
        catch(e) { layer = doc.layers.add(); layer.name = layerName; }
        layer.locked = false; layer.visible = true;

        var activeIndices = [];
        for (var c = 0; c < checkBoxes.length; c++) 
            if (checkBoxes[c].value) activeIndices.push(c);

        if (activeIndices.length === 0) return;

        var hHead = parseFloat(hHeadVal) || 40;
        var hRow = parseFloat(hRowVal) || 20;
        var padding = parseFloat(padVal) || 15;
        var fSize = parseFloat(fSizeVal) || 10;

        var cTxtHead = makeCMYK.apply(null, uiColors.txtHead);
        var cBgHead  = makeCMYK.apply(null, uiColors.bgHead);
        var cTxtCont = makeCMYK.apply(null, uiColors.txtCont);
        var cBgRow1  = makeCMYK.apply(null, uiColors.bgRow1);
        var cBgRow2  = makeCMYK.apply(null, uiColors.bgRow2);
        var cBlack   = makeCMYK(0,0,0,100);

        var tempText = layer.textFrames.add();
        var tRange = tempText.textRange;
        tRange.characterAttributes.size = fSize;
        try { tRange.characterAttributes.textFont = app.textFonts.getByName(fontName); } catch(e){}

        var colWidths = []; var totalW = 0;
        for (var k = 0; k < activeIndices.length; k++) {
            var colIndex = activeIndices[k]; var maxW = 0;
            for (var r = 0; r < data.length; r++) {
                var txt = (data[r] && data[r].length > colIndex) ? data[r][colIndex] : "";
                txt = txt.replace(/^"|"$/g, '');
                tempText.contents = txt;
                if (tempText.width > maxW) maxW = tempText.width;
            }
            colWidths.push(maxW + padding);
            totalW += (maxW + padding);
        }
        tempText.remove(); 

        var abRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;

        var startX = abRect[0] + ((abRect[2] - abRect[0])/2) - (totalW/2);
        var startY = abRect[1] - 50;

        if (isPreview) cleanupPreview();

        var grp = layer.groupItems.add();
        grp.name = isPreview ? "Tableau_PREVIEW_TEMP" : "Tableau CMJN Final";
        if (isPreview) previewGroup = grp;

        var curY = startY;
        for (var r = 0; r < data.length; r++) {
            var isH = (r === 0);
            var lH = isH ? hHead : hRow;
            var cTxt = isH ? cTxtHead : cTxtCont;
            var cBg = isH ? cBgHead : (r%2!==0 ? cBgRow1 : cBgRow2);

            var curX = startX;
            for (var k = 0; k < activeIndices.length; k++) {
                var w = colWidths[k];
                var cell = (data[r] && data[r].length > activeIndices[k]) ? data[r][activeIndices[k]] : "";
                cell = cell.replace(/^"|"$/g, '');
                
                var rect = grp.pathItems.rectangle(curY, curX, w, lH);
                rect.stroked = true; rect.strokeColor = cBlack; rect.strokeWidth = 0.5;
                rect.filled = true; rect.fillColor = cBg;

                if (cell !== "") {
                    var tf = grp.textFrames.add();
                    tf.contents = cell;
                    var ca = tf.textRange.characterAttributes;
                    ca.size = fSize; ca.fillColor = cTxt;
                    try { ca.textFont = app.textFonts.getByName(fontName); } catch(e) {}
                    if (isH && isBold) { ca.strokeWeight = 0.5; ca.strokeColor = cTxt; }

                    tf.top = curY - (lH - tf.height)/2;
                    tf.left = curX + (padding/2);
                }
                curX += w;
            }
            curY -= lH;
        }

    } catch(err) {
        alert("Erreur: " + err.message);
    }
}

main();
