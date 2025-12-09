/*
  Script Adobe Illustrator : Import CSV -> Tableau (V61 Leading Fix)
  - Correction : Force la désactivation de l'interlignage automatique
    pour que la valeur personnalisée soit prise en compte sur les titres.
  - Toutes fonctionnalités incluses.
*/
#targetengine "session_v61"

// --- 1. MOTEUR DE DESSIN ---
var illustratorEngine = function(hexData) {
    
    $.gc(); 

    function fromHex(s){var r="";for(var i=0;i<s.length;i+=2)r+=String.fromCharCode(parseInt(s.substr(i,2),16));return r;}
    function makeCMYK(arr){var c=new CMYKColor();c.cyan=arr[0];c.magenta=arr[1];c.yellow=arr[2];c.black=arr[3];return c;}

    try {
        var params = eval('(' + fromHex(hexData) + ')'); 
        
        if (app.documents.length === 0) return;
        var doc = app.documents[0]; 
        try { doc.selection = null; } catch(e){}

        var mode = params.mode;

        // NETTOYAGE
        try {
            var items = doc.pageItems;
            for (var i = items.length - 1; i >= 0; i--) {
                var n = items[i].name;
                if (n === "Tableau_PREVIEW_TEMP") items[i].remove();
                else if (params.originalName && n === params.originalName) {
                    if (mode === "update") items[i].remove();
                    else if (mode === "preview") items[i].hidden = true;
                }
            }
        } catch(e){}

        // CALQUE
        var layer; try{layer=doc.layers.getByName("Calque_Tableau");}catch(e){layer=doc.layers.add();layer.name="Calque_Tableau";}
        layer.locked=false; layer.visible=true;

        // DATA
        var data = params.data;
        var cols = params.activeCols;
        var aligns = params.colAligns;
        if (!cols || cols.length === 0) return;

        var maxRowsUser = params.geo.maxRows;
        var rowLimit = data.length;
        if (maxRowsUser > 0 && (maxRowsUser + 1) < data.length) rowLimit = maxRowsUser + 1;

        var cHeadT=makeCMYK(params.colors.txtHead); var cHeadB=makeCMYK(params.colors.bgHead);
        var cContT=makeCMYK(params.colors.txtCont); var cRow1=makeCMYK(params.colors.bgRow1);
        var cRow2=makeCMYK(params.colors.bgRow2); 
        var cBlack=makeCMYK([0,0,0,100]); var cWhite=makeCMYK([0,0,0,0]);

        var geo = params.geo;
        var borders = params.borders; 
        var strokeW = (typeof borders.weight === 'number') ? borders.weight : 0.5;

        // FONT LOAD
        var targetFont;
        try { targetFont = app.textFonts.getByName(geo.fName); } 
        catch(e) { targetFont = app.textFonts[0]; }

        // MESURE
        var tempText = layer.textFrames.add();
        var tRange = tempText.textRange;
        tRange.characterAttributes.size = geo.fSize;
        tRange.characterAttributes.textFont = targetFont;

        var colWidths = []; var totalW = 0;
        var scanLimit = (rowLimit > 200) ? 200 : rowLimit;

        // 1. CALCUL LARGEUR
        for (var k=0; k<cols.length; k++) {
            var maxW = 0;
            var startRow = (geo.wrapHead) ? 1 : 0; 
            
            for (var r=startRow; r<scanLimit; r++) {
                var txt = (data[r] && data[r].length > cols[k]) ? String(data[r][cols[k]]).replace(/^"|"$/g,'') : "";
                if(txt!==""){
                    tempText.contents = txt;
                    try { tempText.textRange.characterAttributes.horizontalScale = 100; } catch(e){}
                    if (tempText.width > maxW) maxW = tempText.width;
                }
            }
            var minW = (geo.wrapHead) ? 30 : 10;
            var cw = Math.max(maxW + geo.pad, minW + geo.pad);
            colWidths.push(cw); 
            totalW += cw;
        }

        // 2. RATIO
        var finalColWidths = [];
        if (geo.totalW > 0 && totalW > 0) {
            var ratio = geo.totalW / totalW;
            if (ratio < 0.1) ratio = 1;
            for(var i=0; i<colWidths.length; i++) finalColWidths.push(colWidths[i] * ratio);
        } else {
            finalColWidths = colWidths;
        }

        // 3. TRAITEMENT TITRES (WRAP)
        var processedHeaders = [];
        var CR = String.fromCharCode(13);

        for (var k=0; k<cols.length; k++) {
            var titleTxt = (data[0] && data[0].length > cols[k]) ? String(data[0][cols[k]]).replace(/^"|"$/g,'') : "";
            
            if (geo.wrapHead) {
                var availW = finalColWidths[k] - geo.pad;
                if (availW < 10) availW = 10;

                tempText.contents = titleTxt;
                try { tempText.textRange.characterAttributes.horizontalScale = 100; } catch(e){}
                
                if (tempText.width > availW) {
                    var words = titleTxt.split(" ");
                    var lines = [];
                    var currentLine = words[0];
                    
                    for (var w = 1; w < words.length; w++) {
                        tempText.contents = currentLine + " " + words[w];
                        if (tempText.width > availW) {
                            lines.push(currentLine);
                            currentLine = words[w];
                        } else {
                            currentLine += " " + words[w];
                        }
                    }
                    lines.push(currentLine);
                    processedHeaders.push(lines.join(CR)); 
                } else {
                    processedHeaders.push(titleTxt);
                }
            } else {
                processedHeaders.push(titleTxt);
            }
        }
        tempText.remove();


        // POSITION
        var startX=0, startY=0;
        if (mode==="update" && params.oldPosition) {
            startX=params.oldPosition[0]; startY=params.oldPosition[1];
        } else {
            try { 
                var abRect = doc.artboards[0].artboardRect;
                var finalTotalW = 0; for(var fw=0; fw<finalColWidths.length; fw++) finalTotalW += finalColWidths[fw];
                startX = abRect[0] + ((abRect[2]-abRect[0])/2) - (finalTotalW/2); 
                startY = abRect[1] - 50;
            } catch(e) { startX=0; startY=0; }
        }

        // GROUPE
        var grp = layer.groupItems.add();
        var finalName = "Tableau_ID_" + (new Date().getTime());
        grp.name = (mode==="preview") ? "Tableau_PREVIEW_TEMP" : finalName;

        // TAGS
        if (mode!=="preview") {
            var tag = grp.tags.add(); tag.name = "CSV_SOURCE_DATA";
            var _JSON={stringify:function(v){return v.toSource()},parse:function(t){return eval('('+t+')')}};
            if(typeof JSON!=='undefined' && JSON.stringify) _JSON=JSON;
            tag.value = _JSON.stringify({data:data, headers:params.headers, config:params});
        }

        // DESSIN
        var curY = startY;
        var totalH = 0;

        for (var r=0; r<rowLimit; r++) {
            var isH = (r===0);
            var lH = isH ? geo.hHead : geo.hRow;
            var cTxt = isH ? cHeadT : cContT;
            var cBg = isH ? cHeadB : (r%2!==0 ? cRow1 : cRow2);
            var curX = startX;
            totalH += lH;

            for (var k=0; k<cols.length; k++) {
                var w = finalColWidths[k];
                var cellContent = "";
                if (isH) cellContent = processedHeaders[k];
                else cellContent = (data[r] && data[r].length > cols[k]) ? String(data[r][cols[k]]).replace(/^"|"$/g,'') : "";

                var align = aligns[k] || "left";

                // Fond
                var rect = grp.pathItems.rectangle(curY, curX, w, lH);
                rect.stroked = false; rect.filled = true; rect.fillColor = cBg;

                // Texte
                if (cellContent !== "") {
                    try {
                        var tf = grp.textFrames.add();
                        tf.contents = cellContent.replace(/\\r/g, CR).replace(/\\n/g, CR);
                        var ca = tf.textRange.characterAttributes;
                        ca.size = geo.fSize; ca.fillColor = cTxt;
                        ca.textFont = targetFont; 

                        // --- GESTION TYPO SPECIFIQUE HEADER + FIX LEADING ---
                        if (isH) {
                            if (geo.isBold) { ca.strokeWeight=0.5; ca.strokeColor=cTxt; }
                            
                            // FIX V61 : On force autoLeading à false
                            if (geo.leading > 0) {
                                ca.autoLeading = false; 
                                ca.leading = geo.leading;
                            }
                        }
                        // ----------------------------------------------------
                        
                        // Auto-Scale
                        var doScale = true;
                        if (isH && geo.wrapHead) doScale = false;

                        var availW = w - geo.pad;
                        if (doScale && availW > 5 && tf.width > availW) {
                            var sc = (availW / tf.width) * 100;
                            if (sc < 40) sc = 40; if (sc > 100) sc = 100;
                            ca.horizontalScale = sc;
                        }

                        // Alignement
                        var posX = 0;
                        var textWidth = tf.width;
                        var textHeight = tf.height;

                        if (align === "center") {
                            try { tf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch(e){}
                            posX = curX + (w / 2) - (textWidth / 2); 
                        } else if (align === "right") {
                            try { tf.textRange.paragraphAttributes.justification = Justification.RIGHT; } catch(e){}
                            posX = curX + w - (geo.pad/2) - textWidth;
                        } else { 
                            try { tf.textRange.paragraphAttributes.justification = Justification.LEFT; } catch(e){}
                            posX = curX + (geo.pad/2);
                        }

                        tf.top = curY - (lH - textHeight)/2;
                        tf.left = posX;

                    } catch(et){}
                }
                curX += w;
            }
            curY -= lH;
        }

        // BORDURES
        function drawLine(x1, y1, x2, y2, w, col) {
            var l = grp.pathItems.add();
            l.setEntirePath([[x1, y1], [x2, y2]]);
            l.stroked = true; l.strokeColor = col; l.strokeWidth = w; l.filled = false;
        }

        var finalTotalW = 0; for(var i=0; i<finalColWidths.length; i++) finalTotalW += finalColWidths[i];

        if (borders.vert) {
            var vx = startX;
            var cVertT = cBlack; if (borders.vertHead === 1) cVertT = cWhite;
            for (var k=0; k<cols.length-1; k++) {
                vx += finalColWidths[k];
                var yBotHead = startY - geo.hHead;
                drawLine(vx, startY, vx, yBotHead, strokeW, cVertT);
                if (totalH > geo.hHead) drawLine(vx, yBotHead, vx, startY - totalH, strokeW, cBlack);
            }
        }

        var hy = startY;
        hy -= geo.hHead;
        if (borders.head) drawLine(startX, hy, startX + finalTotalW, hy, strokeW, cBlack);
        else if (borders.horz) drawLine(startX, hy, startX + finalTotalW, hy, strokeW, cBlack);

        if (borders.horz) {
            for (var r=1; r<rowLimit-1; r++) {
                hy -= geo.hRow;
                drawLine(startX, hy, startX + finalTotalW, hy, strokeW, cBlack);
            }
        }

        if (borders.outer) {
            var out = grp.pathItems.rectangle(startY, startX, finalTotalW, totalH);
            out.filled = false; out.stroked = true; out.strokeColor = cBlack; out.strokeWidth = strokeW;
        }

        app.redraw();
    } catch(err) { alert("MOTEUR: " + err.message); }
};


// --- 2. INTERFACE UTILISATEUR ---

if(typeof JSON!=='object'){JSON={};}(function(){'use strict';if(typeof JSON.stringify!=='function'){JSON.stringify=function(v){return v.toSource();};}if(typeof JSON.parse!=='function'){JSON.parse=function(t){return eval('('+t+')');};}}());

var PREFS_FILE = File(Folder.userData + "/Illustrator_Tableau_Prefs.json");
function savePreferences(o) { try{ PREFS_FILE.open("w"); PREFS_FILE.write(JSON.stringify(o)); PREFS_FILE.close(); }catch(e){} }
function loadPreferences() { try{ if(PREFS_FILE.exists){ PREFS_FILE.open("r"); var c=PREFS_FILE.read(); PREFS_FILE.close(); return JSON.parse(c); } }catch(e){} return null; }

function toHex(str) { var res=""; for(var i=0;i<str.length;i++) res+=str.charCodeAt(i).toString(16).toUpperCase(); return res; }
function cmykToScreenRGB(c, m, y, k) { var C=c/100; var M=m/100; var Y=y/100; var K=k/100; return [1-Math.min(1,C*(1-K)+K), 1-Math.min(1,M*(1-K)+K), 1-Math.min(1,Y*(1-K)+K)]; }

var MM_TO_PT = 2.834645;

function addCompactSlider(p,l,v,cb) {
    var g=p.add("group"); g.alignChildren=["left","center"]; g.spacing=2;
    g.add("statictext",undefined,l).preferredSize.width=10;
    var s=g.add("slider",undefined,parseInt(v),0,100); s.preferredSize.width=60;
    var i=g.add("edittext",undefined,v); i.preferredSize.width=35; 
    s.onChanging=function(){i.text=Math.round(this.value);cb();}; i.onChange=function(){var v=parseFloat(this.text)||0;if(v>100)v=100;this.text=v;s.value=v;cb();};
    return {slider:s, input:i};
}

function createCompactColorBlock(parentTab, title, defC, defM, defY, defK) {
    var grpLine = parentTab.add("group");
    grpLine.orientation = "row"; grpLine.alignChildren = ["center", "center"]; 
    grpLine.alignment = ["center", "top"]; grpLine.spacing = 10;
    var lblTitle = grpLine.add("statictext", undefined, title); lblTitle.preferredSize.width = 50; lblTitle.justify = "right";
    var previewBox = grpLine.add("panel"); previewBox.preferredSize = [25, 25]; 
    var grpSliders = grpLine.add("group"); grpSliders.orientation = "column"; grpSliders.spacing = 0; grpSliders.margins = 0;
    var ui = {}; 
    var updatePreview = function() {
        if (!ui.c || !ui.m || !ui.y || !ui.k) return; 
        var rgb = cmykToScreenRGB(ui.c.input.text, ui.m.input.text, ui.y.input.text, ui.k.input.text);
        var g = previewBox.graphics; g.backgroundColor = g.newBrush(g.BrushType.SOLID_COLOR, [rgb[0], rgb[1], rgb[2], 1]);
        try { g.disabledBackgroundColor = g.backgroundColor; } catch(e){} 
        previewBox.visible = false; previewBox.visible = true; 
    };
    ui.c = addCompactSlider(grpSliders, "C", defC, updatePreview);
    ui.m = addCompactSlider(grpSliders, "M", defM, updatePreview);
    ui.y = addCompactSlider(grpSliders, "J", defY, updatePreview);
    ui.k = addCompactSlider(grpSliders, "N", defK, updatePreview);
    updatePreview(); return ui; 
}

function main() {
    var editMode = false; var loadedData = null; var originalGroupName = null; var originalGroupPos = null;
    if (app.documents.length > 0) {
        try {
            var sel = app.documents[0].selection;
            if (sel && sel.length === 1 && sel[0].typename === "GroupItem") {
                var tag = sel[0].tags.getByName("CSV_SOURCE_DATA");
                loadedData = JSON.parse(tag.value); editMode = true;
                originalGroupName = sel[0].name; originalGroupPos = [sel[0].left, sel[0].top];
            }
        } catch(e) {}
    }

    var rawData, headers;
    if (editMode) { rawData = loadedData.data; headers = loadedData.headers; } 
    else {
        if (app.documents.length === 0) { alert("Ouvrez un doc."); return; }
        var f = File.openDialog("CSV?", "*.csv;*.txt;*.*"); if(!f)return;
        f.open("r"); var txt=f.read(); f.close();
        var lines=txt.split(/\r\n|\r|\n/); var sep=(lines[0].indexOf(";")>-1)?";":",";
        rawData=[]; for(var i=0;i<lines.length;i++) if(lines[i]!="") rawData.push(lines[i].split(sep));
        if(rawData.length==0)return; headers=rawData[0];
    }

    var prefs = loadPreferences() || {};
    var cfg = (editMode && loadedData.config) ? loadedData.config : {};
    function getVal(p,d){ try{var x=p.split('.'),v=cfg;for(var i=0;i<x.length;i++)v=v[x[i]];if(v!==undefined)return v;}catch(e){} try{var x=p.split('.'),v=prefs;for(var i=0;i<x.length;i++)v=v[x[i]];if(v!==undefined)return v;}catch(e){} return d;}
    function getCol(k,d){ if(editMode&&cfg.colors&&cfg.colors[k])return cfg.colors[k]; if(prefs.colors&&prefs.colors[k])return prefs.colors[k]; return d;}

    var win = new Window("palette", editMode?"MODIFICATION TABLEAU":"NOUVEAU TABLEAU");
    win.orientation="column"; win.spacing=5; win.margins=10;

    // 1. TYPO HEADER (Full Width)
    var pT = win.add("panel", undefined, "Typographie");
    pT.alignment = "fill"; pT.orientation = "row"; pT.alignChildren = ["left", "center"];
    
    var iF=pT.add("group"); iF.add("statictext",undefined,"Police:"); 
    
    var fontList = []; var sysFonts = app.textFonts;
    // On charge tout
    for(var f=0; f<sysFonts.length; f++) fontList.push(sysFonts[f].name);
    
    var iFont = iF.add("dropdownlist", undefined, fontList); 
    iFont.preferredSize.width=180; 
    
    var savedF = getVal("geo.fName", "ArialMT");
    var selItem = iFont.find(savedF); 
    if(selItem) iFont.selection = selItem; else if(iFont.items.length > 0) iFont.selection = 0;

    var gSize = pT.add("group"); gSize.add("statictext", undefined, "Taille:");
    var sizes = ["6 pt","7 pt","8 pt","9 pt","10 pt","11 pt","12 pt"];
    var iFS = gSize.add("dropdownlist", undefined, sizes); iFS.preferredSize.width = 60;
    var savedSizeVal = getVal("geo.fSize", 10); var sItem = iFS.find(savedSizeVal + " pt");
    if(sItem) iFS.selection = sItem; else iFS.selection = 4;

    // INTERLIGNAGE (LEADING)
    var gLead = pT.add("group"); gLead.add("statictext", undefined, "Interl.:");
    var iLead = gLead.add("edittext", undefined, getVal("geo.leading", 12)); iLead.preferredSize.width = 30;

    var iBold=pT.add("checkbox",undefined,"Gras"); iBold.value=getVal("geo.isBold",true);

    // 2. ZONE CENTRALE
    var grpMid = win.add("group");
    grpMid.orientation = "row";
    grpMid.alignChildren = ["fill", "top"];
    grpMid.spacing = 10;

    // COL 1: Géométrie
    var pG = grpMid.add("panel",undefined,"Géométrie"); 
    pG.orientation = "column"; pG.alignChildren = "left";
    var addInp=function(p,l,v){ var g=p.add("group"); g.add("statictext",undefined,l).preferredSize.width=45; var inp = g.add("edittext",undefined,v); inp.preferredSize.width=35; return inp; };
    
    var valW_Pt = getVal("geo.totalW", 0);
    var displayW = (valW_Pt > 0) ? (Math.round((valW_Pt / MM_TO_PT) * 100) / 100) : 0;
    var iTotalW = addInp(pG,"Larg(mm):", displayW);
    var iHH=addInp(pG,"H.Titre:",getVal("geo.hHead",40)); 
    var iHR=addInp(pG,"H.Ligne:",getVal("geo.hRow",20)); 
    var iPad=addInp(pG,"Marge:",getVal("geo.pad",15));
    var iMaxR=addInp(pG,"Max Lig:",getVal("geo.maxRows",0));

    // OPTION WRAP
    var iWrap = pG.add("checkbox", undefined, "Wrap Titre");
    iWrap.value = getVal("geo.wrapHead", false);

    // COL 2: Bordures
    var pB = grpMid.add("panel", undefined, "Bordures"); 
    pB.orientation = "column"; pB.alignChildren = "left";
    var chkVert = pB.add("checkbox", undefined, "Verticales"); 
    var chkHorz = pB.add("checkbox", undefined, "Horizontales");
    var chkHead = pB.add("checkbox", undefined, "Sous-Titre");
    var chkOut  = pB.add("checkbox", undefined, "Extérieur");
    chkVert.value = getVal("borders.vert", true); chkHorz.value = getVal("borders.horz", true);
    chkHead.value = getVal("borders.head", true); chkOut.value  = getVal("borders.outer", true);
    
    var grpW = pB.add("group"); grpW.alignChildren=["left","center"]; grpW.add("statictext", undefined, "Epais:");
    var weightList = ["0.25 pt", "0.5 pt", "0.75 pt", "1 pt", "2 pt"];
    var dropWeight = grpW.add("dropdownlist", undefined, weightList); dropWeight.preferredSize.width = 60; dropWeight.selection = getVal("borders.weightIdx", 1); 
    
    var grpVH = pB.add("group"); grpVH.alignChildren=["left","center"]; grpVH.add("statictext", undefined, "V. Titre:");
    var vertHeadList = ["Auto", "Blanc", "Noir"];
    var dropVertHead = grpVH.add("dropdownlist", undefined, vertHeadList); dropVertHead.preferredSize.width = 60; dropVertHead.selection = getVal("borders.vertHead", 0);

    // COL 3: Couleurs
    var grpColors = grpMid.add("group");
    grpColors.orientation = "column";
    var tabs = grpColors.add("tabbedpanel"); tabs.preferredSize=[220,230];
    var t1 = tabs.add("tab",undefined,"Entête"); t1.alignChildren=["center","top"];
    var uTH = createCompactColorBlock(t1, "Texte",  getCol("txtHead",[0,0,0,0])[0], getCol("txtHead",[0,0,0,0])[1], getCol("txtHead",[0,0,0,0])[2], getCol("txtHead",[0,0,0,0])[3]);
    var uBH = createCompactColorBlock(t1, "Fond",   getCol("bgHead",[0,0,0,80])[0], getCol("bgHead",[0,0,0,80])[1], getCol("bgHead",[0,0,0,80])[2], getCol("bgHead",[0,0,0,80])[3]);
    var t2 = tabs.add("tab",undefined,"Contenu"); t2.alignChildren=["center","top"];
    var uTC = createCompactColorBlock(t2, "Texte",  getCol("txtCont",[0,0,0,100])[0], getCol("txtCont",[0,0,0,100])[1], getCol("txtCont",[0,0,0,100])[2], getCol("txtCont",[0,0,0,100])[3]);
    var uR1 = createCompactColorBlock(t2, "Ligne 1",getCol("bgRow1",[0,0,0,0])[0], getCol("bgRow1",[0,0,0,0])[1], getCol("bgRow1",[0,0,0,0])[2], getCol("bgRow1",[0,0,0,0])[3]);
    var uR2 = createCompactColorBlock(t2, "Ligne 2",getCol("bgRow2",[0,0,0,10])[0], getCol("bgRow2",[0,0,0,10])[1], getCol("bgRow2",[0,0,0,10])[2], getCol("bgRow2",[0,0,0,10])[3]);
    tabs.selection=0;

    // BAS
    var pCols = win.add("panel", undefined, "Gestion des Colonnes"); pCols.orientation = "column"; pCols.alignChildren = ["center", "top"]; pCols.alignment = ["fill", "top"]; 
    var gDual = pCols.add("group"); gDual.orientation = "row"; gDual.alignChildren = ["center", "center"]; gDual.alignment = ["center", "center"];
    var gLBox = gDual.add("group"); gLBox.orientation="column"; gLBox.alignChildren=["left","top"]; gLBox.spacing=0;
    gLBox.add("statictext",undefined,"Masquées:"); var lstL = gLBox.add("listbox", [0,0,200,150], [], {multiselect: true});
    var gMid = gDual.add("group"); gMid.orientation="column"; gMid.alignment="center";
    var btnRight = gMid.add("button", undefined, "Ajouter >"); btnRight.preferredSize=[70,25];
    var btnLeft = gMid.add("button", undefined, "< Retirer"); btnLeft.preferredSize=[70,25];
    var gRBox = gDual.add("group"); gRBox.orientation="column"; gRBox.alignChildren=["left","top"]; gRBox.spacing=0;
    gRBox.add("statictext",undefined,"Affichées (Ordre):"); var lstR = gRBox.add("listbox", [0,0,200,150], [], {multiselect: true});
    
    var gAlign = gRBox.add("group"); gAlign.orientation="row"; gAlign.spacing=5;
    var bAlL = gAlign.add("button", undefined, "|←"); bAlL.preferredSize=[40,25]; bAlL.helpTip="Gauche";
    var bAlC = gAlign.add("button", undefined, "→|←"); bAlC.preferredSize=[40,25]; bAlC.helpTip="Centre";
    var bAlR = gAlign.add("button", undefined, "→|"); bAlR.preferredSize=[40,25]; bAlR.helpTip="Droite";

    var gSort = gDual.add("group"); gSort.orientation="column"; gSort.alignment="center";
    var btnUp = gSort.add("button", undefined, "▲"); btnUp.preferredSize=[25,25]; var btnDown = gSort.add("button", undefined, "▼"); btnDown.preferredSize=[25,25];

    function populateDualLists() {
        var activeCols = []; var hiddenCols = [];
        if (editMode && cfg.activeCols) {
            for (var k=0; k<cfg.activeCols.length; k++) {
                var idx = cfg.activeCols[k];
                var al = (cfg.colAligns && cfg.colAligns[k]) ? cfg.colAligns[k] : "left";
                activeCols.push({idx: idx, name: headers[idx], align: al});
            }
            for (var h=0; h<headers.length; h++) {
                var isAct = false; for (var x=0; x<cfg.activeCols.length; x++) if(cfg.activeCols[x] == h) isAct=true;
                if (!isAct) hiddenCols.push({idx: h, name: headers[h]});
            }
        } else {
            for (var h=0; h<headers.length; h++) activeCols.push({idx: h, name: headers[h], align: "left"});
        }
        for (var i=0; i<hiddenCols.length; i++) { var it = lstL.add("item", hiddenCols[i].name); it.realIndex = hiddenCols[i].idx; }
        for (var j=0; j<activeCols.length; j++) { 
            var displayTxt = getAlignSymbol(activeCols[j].align) + " " + activeCols[j].name;
            var it2 = lstR.add("item", displayTxt); it2.realIndex = activeCols[j].idx; it2.alignmentVal = activeCols[j].align;
        }
    }
    function getAlignSymbol(a) { if(a=="center") return "[C]"; if(a=="right") return "[R]"; return "[L]"; }

    populateDualLists();

    function setAlign(val) {
        var sel = lstR.selection; if(!sel) return; if(!(sel instanceof Array)) sel = [sel];
        for(var i=0; i<sel.length; i++) {
            sel[i].alignmentVal = val;
            sel[i].text = getAlignSymbol(val) + " " + headers[sel[i].realIndex];
        }
    }
    bAlL.onClick = function(){ setAlign("left"); };
    bAlC.onClick = function(){ setAlign("center"); };
    bAlR.onClick = function(){ setAlign("right"); };

    function moveItems(source, dest) {
        var sels = source.selection; if (!sels) return; if (!(sels instanceof Array)) sels = [sels];
        for (var i=0; i<sels.length; i++) { 
            var item = sels[i]; 
            var txt = (source == lstL) ? "[L] " + item.text : item.text;
            if (source == lstR) txt = headers[item.realIndex];
            var newItem = dest.add("item", txt); newItem.realIndex = item.realIndex; 
            if(source == lstL) newItem.alignmentVal = "left";
        }
        var toRemove = []; for(var k=0; k<source.items.length; k++) if(source.items[k].selected) toRemove.push(k);
        for(var k=toRemove.length-1; k>=0; k--) source.remove(toRemove[k]);
    }
    btnRight.onClick = function() { moveItems(lstL, lstR); }; btnLeft.onClick = function() { moveItems(lstR, lstL); };
    function swap(i1, i2) {
        var t1=lstR.items[i1].text, r1=lstR.items[i1].realIndex, s1=lstR.items[i1].selected, a1=lstR.items[i1].alignmentVal;
        var t2=lstR.items[i2].text, r2=lstR.items[i2].realIndex, s2=lstR.items[i2].selected, a2=lstR.items[i2].alignmentVal;
        lstR.items[i1].text=t2; lstR.items[i1].realIndex=r2; lstR.items[i1].selected=s2; lstR.items[i1].alignmentVal=a2;
        lstR.items[i2].text=t1; lstR.items[i2].realIndex=r1; lstR.items[i2].selected=s1; lstR.items[i2].alignmentVal=a1;
    }
    btnUp.onClick = function() { var s = lstR.selection; if(!s)return; for(var i=1; i<lstR.items.length; i++) if(lstR.items[i].selected && !lstR.items[i-1].selected) swap(i,i-1); };
    btnDown.onClick = function() { var s = lstR.selection; if(!s)return; for(var i=lstR.items.length-2; i>=0; i--) if(lstR.items[i].selected && !lstR.items[i+1].selected) swap(i,i+1); };

    // BTNS
    var gB = win.add("group");
    var bPrev = gB.add("button",undefined,"Prévisualiser"); bPrev.preferredSize.width = 90;
    var bSave = gB.add("button",undefined,"Sauver Préf."); bSave.preferredSize.width = 80;
    var bX = gB.add("button",undefined,"Annuler"); bX.preferredSize.width = 70;
    var bOK = gB.add("button",undefined,"Valider"); bOK.preferredSize.width = 70;

    function collectParams(isPreview) {
        var act = []; var aligns = [];
        for(var i=0; i<lstR.items.length; i++) {
            act.push(lstR.items[i].realIndex);
            aligns.push(lstR.items[i].alignmentVal || "left");
        }
        function rc(u) { return [parseFloat(u.c.input.text), parseFloat(u.m.input.text), parseFloat(u.y.input.text), parseFloat(u.k.input.text)]; }
        var wStr = dropWeight.selection.text.replace(" pt", ""); var wVal = parseFloat(wStr) || 0.5;
        var widthMm = parseFloat(iTotalW.text) || 0; var widthPt = widthMm * MM_TO_PT;
        var maxRowsVal = parseFloat(iMaxR.text) || 0;
        
        var fSizeStr = "10";
        if(iFS.selection) fSizeStr = iFS.selection.text.replace(" pt", "");
        var fSizeVal = parseFloat(fSizeStr) || 10;
        var fNameVal = (iFont.selection) ? iFont.selection.text : "ArialMT";

        return {
            data: rawData, headers: headers, activeCols: act, colAligns: aligns,
            isPreview: isPreview, mode: isPreview ? "preview" : (editMode ? "update" : "create"),
            originalName: originalGroupName, oldPosition: originalGroupPos,
            geo: { 
                hHead: parseFloat(iHH.text)||40, hRow: parseFloat(iHR.text)||20, pad: parseFloat(iPad.text)||15, 
                fSize: fSizeVal, fName: fNameVal, isBold: iBold.value,
                totalW: widthPt, maxRows: maxRowsVal, wrapHead: iWrap.value,
                leading: parseFloat(iLead.text) || 12
            },
            colors: { txtHead: rc(uTH), bgHead: rc(uBH), txtCont: rc(uTC), bgRow1: rc(uR1), bgRow2: rc(uR2) },
            borders: { vert: chkVert.value, horz: chkHorz.value, head: chkHead.value, outer: chkOut.value, weight: wVal, weightIdx: dropWeight.selection.index, vertHead: dropVertHead.selection.index }
        };
    }

    function send(params) {
        $.sleep(100); 
        var jsonStr = JSON.stringify(params); var hexStr = toHex(jsonStr);
        var bt = new BridgeTalk(); bt.target = "illustrator";
        bt.body = "(" + illustratorEngine.toString() + ")('" + hexStr + "');";
        bt.send();
    }

    bPrev.onClick = function() { send(collectParams(true)); };
    bOK.onClick = function() { send(collectParams(false)); win.close(); };
    bSave.onClick = function() { var p = collectParams(true); savePreferences({ geo: p.geo, colors: p.colors, borders: p.borders }); alert("Sauvegardé!"); };
    bX.onClick = function() { 
        var bt = new BridgeTalk(); bt.target = "illustrator"; 
        bt.body = "try{var d=app.activeDocument;var i=d.pageItems;for(var j=i.length-1;j>=0;j--)if(i[j].name=='Tableau_PREVIEW_TEMP')i[j].remove();}catch(e){}"; 
        bt.send(); win.close(); 
    };

    win.show();
}

main();