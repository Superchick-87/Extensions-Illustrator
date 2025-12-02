/*
  Script Adobe Illustrator : Import CSV -> Tableau (V42 Perfect Visibility Fix)
  - Réintégration stricte de la gestion de visibilité (Masquer/Afficher/Supprimer) pour l'ancien tableau.
  - Toutes options incluses (Bordures titre, Epaisseur, UI 35px, etc.)
*/
#targetengine "session_v42"

// --- 1. MOTEUR DE DESSIN ---
var illustratorEngine = function(hexData) {
    function fromHex(s){var r="";for(var i=0;i<s.length;i+=2)r+=String.fromCharCode(parseInt(s.substr(i,2),16));return r;}
    function makeCMYK(arr){var c=new CMYKColor();c.cyan=arr[0];c.magenta=arr[1];c.yellow=arr[2];c.black=arr[3];return c;}

    try {
        var params = eval('(' + fromHex(hexData) + ')'); 
        
        if (app.documents.length === 0) return;
        var doc = app.documents[0];
        doc.activate(); 
        doc.selection = null; 

        var mode = params.mode;

        // --- GESTION VISIBILITE CRITIQUE ---
        try {
            var items = doc.pageItems;
            for (var i = items.length - 1; i >= 0; i--) {
                var n = items[i].name;
                
                // 1. Toujours nettoyer les previews qui trainent
                if (n === "Tableau_PREVIEW_TEMP") {
                    items[i].remove();
                }
                // 2. Gestion de l'original (si mode modification)
                else if (params.originalName && n === params.originalName) {
                    if (mode === "update") {
                        // Si validation finale : on supprime l'ancien
                        items[i].remove();
                    } else if (mode === "preview") {
                        // Si preview : on le cache seulement (pour ne pas voir de doublon)
                        items[i].hidden = true;
                    }
                }
            }
        } catch(e){}
        // -----------------------------------

        // CALQUE
        var layer; try{layer=doc.layers.getByName("Calque_Tableau");}catch(e){layer=doc.layers.add();layer.name="Calque_Tableau";}
        layer.locked=false; layer.visible=true;

        // DATA
        var data = params.data;
        var cols = params.activeCols;
        if (!cols || cols.length === 0) return;

        var cHeadT=makeCMYK(params.colors.txtHead); var cHeadB=makeCMYK(params.colors.bgHead);
        var cContT=makeCMYK(params.colors.txtCont); var cRow1=makeCMYK(params.colors.bgRow1);
        var cRow2=makeCMYK(params.colors.bgRow2); 
        var cBlack=makeCMYK([0,0,0,100]);
        var cWhite=makeCMYK([0,0,0,0]);

        var geo = params.geo;
        var borders = params.borders; 
        var strokeW = (borders.weight !== undefined) ? borders.weight : 0.5;

        // MESURE
        var tempText = layer.textFrames.add();
        var tRange = tempText.textRange;
        tRange.characterAttributes.size = geo.fSize;
        try { tRange.characterAttributes.textFont = app.textFonts.getByName(geo.fName); } catch(e){}

        var colWidths = []; var totalW = 0;
        for (var k=0; k<cols.length; k++) {
            var maxW = 0;
            for (var r=0; r<data.length; r++) {
                var txt = (data[r] && data[r].length > cols[k]) ? String(data[r][cols[k]]).replace(/^"|"$/g,'') : "";
                tempText.contents = txt;
                if (tempText.width > maxW) maxW = tempText.width;
            }
            colWidths.push(maxW + geo.pad); totalW += (maxW + geo.pad);
        }
        tempText.remove();

        // POSITION
        var startX=0, startY=0;
        if (mode==="update" && params.oldPosition) {
            startX=params.oldPosition[0]; startY=params.oldPosition[1];
        } else {
            try { 
                var abRect = doc.artboards[0].artboardRect;
                startX = abRect[0] + ((abRect[2]-abRect[0])/2) - (totalW/2); 
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

        // PHASE 1 : FONDS & TEXTE
        var curY = startY;
        var totalH = 0;

        for (var r=0; r<data.length; r++) {
            var isH = (r===0);
            var lH = isH ? geo.hHead : geo.hRow;
            var cTxt = isH ? cHeadT : cContT;
            var cBg = isH ? cHeadB : (r%2!==0 ? cRow1 : cRow2);
            var curX = startX;
            totalH += lH;

            for (var k=0; k<cols.length; k++) {
                var w = colWidths[k];
                var cell = (data[r] && data[r].length > cols[k]) ? String(data[r][cols[k]]).replace(/^"|"$/g,'') : "";
                
                var rect = grp.pathItems.rectangle(curY, curX, w, lH);
                rect.stroked = false;
                rect.filled = true; 
                rect.fillColor = cBg;

                if (cell !== "") {
                    var tf = grp.textFrames.add();
                    tf.contents = cell;
                    var ca = tf.textRange.characterAttributes;
                    ca.size = geo.fSize; ca.fillColor = cTxt;
                    try { ca.textFont = app.textFonts.getByName(geo.fName); } catch(e) {}
                    if (isH && geo.isBold) { ca.strokeWeight=0.5; ca.strokeColor=cTxt; }
                    try { tf.textRange.paragraphAttributes.justification = Justification.LEFT; } catch(e){}
                    tf.top = curY - (lH - tf.height)/2;
                    tf.left = curX + (geo.pad/2);
                }
                curX += w;
            }
            curY -= lH;
        }

        // PHASE 2 : BORDURES
        function drawLine(x1, y1, x2, y2, weight, color) {
            var line = grp.pathItems.add();
            line.setEntirePath([[x1, y1], [x2, y2]]);
            line.stroked = true;
            line.strokeColor = color;
            line.strokeWidth = weight;
            line.filled = false;
        }

        // 1. Lignes Verticales
        if (borders.vert) {
            var vx = startX;
            
            var cVertHead = cBlack; 
            if (borders.vertHead === 1) cVertHead = cWhite; 
            if (borders.vertHead === 2) cVertHead = cBlack; 

            for (var k=0; k<cols.length-1; k++) {
                vx += colWidths[k];
                var yHeadBot = startY - geo.hHead;
                drawLine(vx, startY, vx, yHeadBot, strokeW, cVertHead);
                if (totalH > geo.hHead) {
                    drawLine(vx, yHeadBot, vx, startY - totalH, strokeW, cBlack);
                }
            }
        }

        var hy = startY;
        hy -= geo.hHead; 
        
        // 2. Sous-titre
        if (borders.head) {
            drawLine(startX, hy, startX + totalW, hy, strokeW, cBlack);
        } else if (borders.horz) {
            drawLine(startX, hy, startX + totalW, hy, strokeW, cBlack);
        }

        // 3. Autres lignes
        if (borders.horz) {
            for (var r=1; r<data.length-1; r++) {
                hy -= geo.hRow;
                drawLine(startX, hy, startX + totalW, hy, strokeW, cBlack);
            }
        }

        // 4. Contour Extérieur
        if (borders.outer) {
            var border = grp.pathItems.rectangle(startY, startX, totalW, totalH);
            border.filled = false;
            border.stroked = true;
            border.strokeColor = cBlack;
            border.strokeWidth = strokeW;
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

function cmykToScreenRGB(c, m, y, k) {
    var C = c/100; var M = m/100; var Y = y/100; var K = k/100;
    var r = 1 - Math.min(1, C * (1 - K) + K);
    var g = 1 - Math.min(1, M * (1 - K) + K);
    var b = 1 - Math.min(1, Y * (1 - K) + K);
    return [r, g, b];
}

function addCompactSlider(p,l,v,cb) {
    var g=p.add("group"); g.alignChildren=["left","center"]; g.spacing=2;
    g.add("statictext",undefined,l).preferredSize.width=10;
    var s=g.add("slider",undefined,parseInt(v),0,100); s.preferredSize.width=60;
    var i=g.add("edittext",undefined,v); 
    i.preferredSize.width=35; 
    s.onChanging=function(){i.text=Math.round(this.value);cb();}; i.onChange=function(){var v=parseFloat(this.text)||0;if(v>100)v=100;this.text=v;s.value=v;cb();};
    return {slider:s, input:i};
}

function createCompactColorBlock(parentTab, title, defC, defM, defY, defK) {
    var grpLine = parentTab.add("group");
    grpLine.orientation = "row"; grpLine.alignChildren = ["center", "center"]; 
    grpLine.alignment = ["center", "top"];
    grpLine.spacing = 10;

    var lblTitle = grpLine.add("statictext", undefined, title);
    lblTitle.preferredSize.width = 50; 
    lblTitle.justify = "right";

    var previewBox = grpLine.add("panel"); 
    previewBox.preferredSize = [25, 25]; 

    var grpSliders = grpLine.add("group"); 
    grpSliders.orientation = "column"; 
    grpSliders.spacing = 0; 
    grpSliders.margins = 0;

    var ui = {}; 
    var updatePreview = function() {
        if (!ui.c || !ui.m || !ui.y || !ui.k) return; 
        var rgb = cmykToScreenRGB(ui.c.input.text, ui.m.input.text, ui.y.input.text, ui.k.input.text);
        var g = previewBox.graphics;
        g.backgroundColor = g.newBrush(g.BrushType.SOLID_COLOR, [rgb[0], rgb[1], rgb[2], 1]);
        try { g.disabledBackgroundColor = g.backgroundColor; } catch(e){} 
        previewBox.visible = false; previewBox.visible = true; 
    };

    ui.c = addCompactSlider(grpSliders, "C", defC, updatePreview);
    ui.m = addCompactSlider(grpSliders, "M", defM, updatePreview);
    ui.y = addCompactSlider(grpSliders, "J", defY, updatePreview);
    ui.k = addCompactSlider(grpSliders, "N", defK, updatePreview);
    updatePreview();
    return ui; 
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

    // WINDOW
    var win = new Window("palette", editMode?"MODIFICATION TABLEAU":"NOUVEAU TABLEAU");
    win.orientation="column"; win.spacing=5; win.margins=10;

    var grpTop = win.add("group"); grpTop.orientation="row"; grpTop.alignChildren=["center","top"]; grpTop.spacing=10;

    var c1 = grpTop.add("group"); c1.orientation="column"; c1.spacing=5;

    // Panel Geo
    var pG = c1.add("panel",undefined,"Géométrie"); pG.alignChildren="left";
    var addInp=function(p,l,v){
        var g=p.add("group");
        g.add("statictext",undefined,l).preferredSize.width=45;
        var inp = g.add("edittext",undefined,v);
        inp.preferredSize.width=35; 
        return inp;
    };
    var iHH=addInp(pG,"H.Titre:",getVal("geo.hHead",40)); 
    var iHR=addInp(pG,"H.Ligne:",getVal("geo.hRow",20)); 
    var iPad=addInp(pG,"Marge:",getVal("geo.pad",15));

    // Panel Typo
    var pT = c1.add("panel",undefined,"Typo"); pT.alignChildren="left";
    var iF=pT.add("group"); iF.add("statictext",undefined,"Font:"); var iFont=iF.add("edittext",undefined,getVal("geo.fName","ArialMT")); iFont.preferredSize.width=60;
    var iFS=addInp(pT,"Taille:",getVal("geo.fSize",10)); 
    var iBold=pT.add("checkbox",undefined,"Gras (Titre)"); iBold.value=getVal("geo.isBold",true);

    // Panel Bordures
    var c2 = grpTop.add("group"); c2.orientation="column"; c2.spacing=5;
    var pB = c2.add("panel", undefined, "Bordures"); pB.alignChildren="left"; pB.preferredSize=[140, -1];
    
    var chkVert = pB.add("checkbox", undefined, "Verticales"); 
    var chkHorz = pB.add("checkbox", undefined, "Horizontales");
    var chkHead = pB.add("checkbox", undefined, "Sous-Titre");
    var chkOut  = pB.add("checkbox", undefined, "Extérieur");
    chkVert.value = getVal("borders.vert", true);
    chkHorz.value = getVal("borders.horz", true);
    chkHead.value = getVal("borders.head", true);
    chkOut.value  = getVal("borders.outer", true);

    var grpW = pB.add("group"); grpW.alignChildren=["left","center"];
    grpW.add("statictext", undefined, "Épaisseur:");
    var weightList = ["0.25 pt", "0.5 pt", "0.75 pt", "1 pt", "2 pt"];
    var dropWeight = grpW.add("dropdownlist", undefined, weightList);
    dropWeight.preferredSize.width = 60;
    dropWeight.selection = getVal("borders.weightIdx", 1); 

    var grpVH = pB.add("group"); grpVH.alignChildren=["left","center"];
    grpVH.add("statictext", undefined, "Vert. Titre:");
    var vertHeadList = ["Auto (Noir)", "Blanc", "Noir"];
    var dropVertHead = grpVH.add("dropdownlist", undefined, vertHeadList);
    dropVertHead.preferredSize.width = 60;
    dropVertHead.selection = getVal("borders.vertHead", 0);

    // Couleurs
    var c3 = grpTop.add("group"); c3.orientation="column";
    var tabs = c3.add("tabbedpanel"); tabs.preferredSize=[230,230];
    
    var t1 = tabs.add("tab",undefined,"Entête"); t1.alignChildren=["center","top"];
    var uTH = createCompactColorBlock(t1, "Texte",  getCol("txtHead",[0,0,0,0])[0], getCol("txtHead",[0,0,0,0])[1], getCol("txtHead",[0,0,0,0])[2], getCol("txtHead",[0,0,0,0])[3]);
    var uBH = createCompactColorBlock(t1, "Fond",   getCol("bgHead",[0,0,0,80])[0], getCol("bgHead",[0,0,0,80])[1], getCol("bgHead",[0,0,0,80])[2], getCol("bgHead",[0,0,0,80])[3]);
    
    var t2 = tabs.add("tab",undefined,"Contenu"); t2.alignChildren=["center","top"];
    var uTC = createCompactColorBlock(t2, "Texte",  getCol("txtCont",[0,0,0,100])[0], getCol("txtCont",[0,0,0,100])[1], getCol("txtCont",[0,0,0,100])[2], getCol("txtCont",[0,0,0,100])[3]);
    var uR1 = createCompactColorBlock(t2, "Ligne 1",getCol("bgRow1",[0,0,0,0])[0], getCol("bgRow1",[0,0,0,0])[1], getCol("bgRow1",[0,0,0,0])[2], getCol("bgRow1",[0,0,0,0])[3]);
    var uR2 = createCompactColorBlock(t2, "Ligne 2",getCol("bgRow2",[0,0,0,10])[0], getCol("bgRow2",[0,0,0,10])[1], getCol("bgRow2",[0,0,0,10])[2], getCol("bgRow2",[0,0,0,10])[3]);
    tabs.selection=0;

    // --- ZONE BASSE ---
    var pCols = win.add("panel", undefined, "Gestion des Colonnes"); 
    pCols.orientation = "column"; pCols.alignChildren = ["center", "top"]; pCols.alignment = ["fill", "top"]; 
    var gDual = pCols.add("group"); gDual.orientation = "row"; gDual.alignChildren = ["center", "center"]; gDual.alignment = ["center", "center"];
    var gLBox = gDual.add("group"); gLBox.orientation="column"; gLBox.alignChildren=["left","top"]; gLBox.spacing=0;
    gLBox.add("statictext",undefined,"Masquées:");
    var lstL = gLBox.add("listbox", [0,0,200,150], [], {multiselect: true});
    var gMid = gDual.add("group"); gMid.orientation="column"; gMid.alignment="center";
    var btnRight = gMid.add("button", undefined, "Ajouter >"); btnRight.preferredSize=[70,25];
    var btnLeft = gMid.add("button", undefined, "< Retirer"); btnLeft.preferredSize=[70,25];
    var gRBox = gDual.add("group"); gRBox.orientation="column"; gRBox.alignChildren=["left","top"]; gRBox.spacing=0;
    gRBox.add("statictext",undefined,"Affichées (Ordre):");
    var lstR = gRBox.add("listbox", [0,0,200,150], [], {multiselect: true});
    var gSort = gDual.add("group"); gSort.orientation="column"; gSort.alignment="center";
    var btnUp = gSort.add("button", undefined, "↑"); btnUp.preferredSize=[25,25];
    var btnDown = gSort.add("button", undefined, "↓"); btnDown.preferredSize=[25,25];

    function populateDualLists() {
        var activeCols = []; var hiddenCols = [];
        if (editMode && cfg.activeCols) {
            for (var k=0; k<cfg.activeCols.length; k++) activeCols.push({idx: cfg.activeCols[k], name: headers[cfg.activeCols[k]]});
            for (var h=0; h<headers.length; h++) {
                var isAct = false; for (var x=0; x<cfg.activeCols.length; x++) if(cfg.activeCols[x] == h) isAct=true;
                if (!isAct) hiddenCols.push({idx: h, name: headers[h]});
            }
        } else {
            for (var h=0; h<headers.length; h++) activeCols.push({idx: h, name: headers[h]});
        }
        for (var i=0; i<hiddenCols.length; i++) { var it = lstL.add("item", hiddenCols[i].name); it.realIndex = hiddenCols[i].idx; }
        for (var j=0; j<activeCols.length; j++) { var it2 = lstR.add("item", activeCols[j].name); it2.realIndex = activeCols[j].idx; }
    }
    populateDualLists();

    function moveItems(source, dest) {
        var sels = source.selection; if (!sels) return; if (!(sels instanceof Array)) sels = [sels];
        for (var i=0; i<sels.length; i++) { var item = sels[i]; var newItem = dest.add("item", item.text); newItem.realIndex = item.realIndex; }
        var toRemove = []; for(var k=0; k<source.items.length; k++) if(source.items[k].selected) toRemove.push(k);
        for(var k=toRemove.length-1; k>=0; k--) source.remove(toRemove[k]);
    }
    btnRight.onClick = function() { moveItems(lstL, lstR); };
    btnLeft.onClick = function() { moveItems(lstR, lstL); };

    function swap(i1, i2) {
        var t1=lstR.items[i1].text, r1=lstR.items[i1].realIndex, s1=lstR.items[i1].selected;
        var t2=lstR.items[i2].text, r2=lstR.items[i2].realIndex, s2=lstR.items[i2].selected;
        lstR.items[i1].text=t2; lstR.items[i1].realIndex=r2; lstR.items[i1].selected=s2;
        lstR.items[i2].text=t1; lstR.items[i2].realIndex=r1; lstR.items[i2].selected=s1;
    }
    btnUp.onClick = function() { var s = lstR.selection; if(!s)return; for(var i=1; i<lstR.items.length; i++) if(lstR.items[i].selected && !lstR.items[i-1].selected) swap(i,i-1); };
    btnDown.onClick = function() { var s = lstR.selection; if(!s)return; for(var i=lstR.items.length-2; i>=0; i--) if(lstR.items[i].selected && !lstR.items[i+1].selected) swap(i,i+1); };

    // BTNS
    var gB = win.add("group");
    var bPrev = gB.add("button",undefined,"Preview");
    var bOK = gB.add("button",undefined, editMode?"Update":"Créer");
    var bSave = gB.add("button",undefined,"Sauver Préf.");
    var bX = gB.add("button",undefined,"Fermer");

    function collectParams(isPreview) {
        var act = []; for(var i=0; i<lstR.items.length; i++) act.push(lstR.items[i].realIndex);
        function rc(u) { return [parseFloat(u.c.input.text), parseFloat(u.m.input.text), parseFloat(u.y.input.text), parseFloat(u.k.input.text)]; }
        
        var wStr = dropWeight.selection.text.replace(" pt", "");
        var wVal = parseFloat(wStr) || 0.5;

        return {
            data: rawData, headers: headers, activeCols: act, isPreview: isPreview,
            mode: isPreview ? "preview" : (editMode ? "update" : "create"),
            originalName: originalGroupName, oldPosition: originalGroupPos,
            geo: { hHead: parseFloat(iHH.text), hRow: parseFloat(iHR.text), pad: parseFloat(iPad.text), fSize: parseFloat(iFS.text), fName: iFont.text, isBold: iBold.value },
            colors: { txtHead: rc(uTH), bgHead: rc(uBH), txtCont: rc(uTC), bgRow1: rc(uR1), bgRow2: rc(uR2) },
            borders: { vert: chkVert.value, horz: chkHorz.value, head: chkHead.value, outer: chkOut.value, weight: wVal, weightIdx: dropWeight.selection.index, vertHead: dropVertHead.selection.index }
        };
    }

    function send(params) {
        var jsonStr = JSON.stringify(params);
        var hexStr = toHex(jsonStr);
        var bt = new BridgeTalk(); bt.target = "illustrator";
        bt.body = "(" + illustratorEngine.toString() + ")('" + hexStr + "');";
        bt.send();
    }

    bPrev.onClick = function() { send(collectParams(true)); };
    bOK.onClick = function() { send(collectParams(false)); win.close(); };
    bSave.onClick = function() { var p = collectParams(true); savePreferences({ geo: p.geo, colors: p.colors, borders: p.borders }); alert("Sauvegardé!"); };
    
    // RESTAURATION SI FERMETURE
    bX.onClick = function() {
        if (editMode && originalGroupName) {
            // Restore original on close
            var restoreScript = "try{var d=app.documents[0];var i=d.pageItems;for(var j=i.length-1;j>=0;j--){if(i[j].name=='Tableau_PREVIEW_TEMP')i[j].remove();if(i[j].name=='" + originalGroupName + "')i[j].hidden=false;}}catch(e){}";
            var bt = new BridgeTalk(); bt.target = "illustrator"; bt.body = restoreScript; bt.send();
        } else {
            var bt = new BridgeTalk(); bt.target = "illustrator";
            bt.body = "try{var d=app.documents[0];var i=d.pageItems;for(var j=i.length-1;j>=0;j--)if(i[j].name=='Tableau_PREVIEW_TEMP')i[j].remove();}catch(e){}";
            bt.send();
        }
        win.close(); 
    };

    win.show();
}

main();