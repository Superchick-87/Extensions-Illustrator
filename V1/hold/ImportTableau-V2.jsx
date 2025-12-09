/*
  Script Adobe Illustrator : Import CSV -> Tableau (V14 Architecture BridgeTalk)
  Cette version sépare totalement l'interface du moteur de dessin.
  Plus aucune erreur "No Document" n'est possible car le dessin est délégué.
*/
#targetengine "session_v14_ui"

// --- PARTIE 1 : LE MOTEUR DE DESSIN (Ce code est envoyé à Illustrator) ---
// Cette fonction contient TOUT ce qu'il faut pour dessiner. Elle est autonome.
var illustratorEngine = function(params) {
    
    // Fonctions internes au moteur
    function makeCMYK(c, m, y, k) {
        var color = new CMYKColor();
        color.cyan = c; color.magenta = m; color.yellow = y; color.black = k;
        return color;
    }

    function cleanup(doc) {
        try {
            var items = doc.pageItems;
            for (var i = items.length - 1; i >= 0; i--) {
                if (items[i].name === "Tableau_PREVIEW_TEMP") {
                    items[i].remove();
                }
            }
        } catch(e) {}
    }

    // --- EXECUTION DU DESSIN ---
    try {
        // 1. Accès Document (Garanti car exécuté par Illustrator Main)
        var doc = app.activeDocument; // Ici, ça marche TOUJOURS
        
        // 2. Nettoyage si preview
        if (params.isPreview) cleanup(doc);

        // 3. Préparation Calque
        var layerName = "Calque_Tableau";
        var layer;
        try { layer = doc.layers.getByName(layerName); } 
        catch(e) { layer = doc.layers.add(); layer.name = layerName; }
        layer.locked = false; layer.visible = true;

        // 4. Données
        var data = params.data;
        var cols = params.activeCols;
        if (cols.length === 0) return;

        // 5. Couleurs
        var cTxtHead = makeCMYK.apply(null, params.colors.txtHead);
        var cBgHead  = makeCMYK.apply(null, params.colors.bgHead);
        var cTxtCont = makeCMYK.apply(null, params.colors.txtCont);
        var cBgRow1  = makeCMYK.apply(null, params.colors.bgRow1);
        var cBgRow2  = makeCMYK.apply(null, params.colors.bgRow2);
        var cBlack   = makeCMYK(0,0,0,100);

        // 6. Outil de mesure
        var tempText = layer.textFrames.add();
        var tRange = tempText.textRange;
        tRange.characterAttributes.size = params.geo.fSize;
        try { tRange.characterAttributes.textFont = app.textFonts.getByName(params.geo.fName); } catch(e){}

        var colWidths = []; var totalW = 0;
        for (var k = 0; k < cols.length; k++) {
            var colIndex = cols[k]; var maxW = 0;
            for (var r = 0; r < data.length; r++) {
                var txt = (data[r] && data[r].length > colIndex) ? data[r][colIndex] : "";
                txt = String(txt).replace(/^"|"$/g, '');
                tempText.contents = txt;
                if (tempText.width > maxW) maxW = tempText.width;
            }
            colWidths.push(maxW + params.geo.pad);
            totalW += (maxW + params.geo.pad);
        }
        tempText.remove();

        // 7. Position
        var abRect;
        try { abRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; } 
        catch(e) { abRect = doc.artboards[0].artboardRect; }
        var startX = abRect[0] + ((abRect[2] - abRect[0])/2) - (totalW/2);
        var startY = abRect[1] - 50;

        // 8. Groupe
        var grp = layer.groupItems.add();
        grp.name = params.isPreview ? "Tableau_PREVIEW_TEMP" : "Tableau CMJN Final";

        // 9. Boucle Dessin
        var curY = startY;
        for (var r = 0; r < data.length; r++) {
            var isH = (r === 0);
            var lH = isH ? params.geo.hHead : params.geo.hRow;
            var cTxt = isH ? cTxtHead : cTxtCont;
            var cBg = isH ? cBgHead : (r%2!==0 ? cBgRow1 : cBgRow2);

            var curX = startX;
            for (var k = 0; k < cols.length; k++) {
                var w = colWidths[k];
                var originalIdx = cols[k];
                var cell = (data[r] && data[r].length > originalIdx) ? data[r][originalIdx] : "";
                cell = String(cell).replace(/^"|"$/g, '');
                
                var rect = grp.pathItems.rectangle(curY, curX, w, lH);
                rect.stroked = true; rect.strokeColor = cBlack; rect.strokeWidth = 0.5;
                rect.filled = true; rect.fillColor = cBg;

                if (cell !== "") {
                    var tf = grp.textFrames.add();
                    tf.contents = cell;
                    var ca = tf.textRange.characterAttributes;
                    ca.size = params.geo.fSize; ca.fillColor = cTxt;
                    try { ca.textFont = app.textFonts.getByName(params.geo.fName); } catch(e) {}
                    if (isH && params.geo.isBold) { ca.strokeWeight = 0.5; ca.strokeColor = cTxt; }

                    tf.top = curY - (lH - tf.height)/2;
                    tf.left = curX + (params.geo.pad/2);
                }
                curX += w;
            }
            curY -= lH;
        }
        app.redraw(); // Rafraichissement immédiat

    } catch(err) {
        alert("Erreur Moteur : " + err.message);
    }
};


// --- PARTIE 2 : L'INTERFACE (PALETTE) ---
// Ce code tourne dans la palette et ne touche JAMAIS au document.

function cmykToScreenRGB(c, m, y, k) {
    var C = c/100; var M = m/100; var Y = y/100; var K = k/100;
    var r = 1 - Math.min(1, C*(1-K)+K);
    var g = 1 - Math.min(1, M*(1-K)+K);
    var b = 1 - Math.min(1, Y*(1-K)+K);
    return [r, g, b];
}

function addCompactSlider(parent, label, initVal, callback) {
    var g = parent.add("group");
    g.orientation = "row"; g.alignChildren = ["left", "center"]; g.spacing = 2; g.margins = 0;
    var lbl = g.add("statictext", undefined, label); lbl.preferredSize.width = 10;
    var sld = g.add("slider", undefined, parseInt(initVal), 0, 100); sld.preferredSize.width = 60;
    var inp = g.add("edittext", undefined, initVal); inp.preferredSize.width = 25;
    sld.onChanging = function() { inp.text = Math.round(this.value); callback(); };
    inp.onChange = function() {
        var v = parseFloat(this.text); if(isNaN(v)) v=0; if(v<0) v=0; if(v>100) v=100;
        this.text = v; sld.value = v; callback();
    };
    return { slider: sld, input: inp };
}

function createCompactColorBlock(parentTab, title, defC, defM, defY, defK) {
    var grpLine = parentTab.add("group");
    grpLine.orientation = "row"; grpLine.alignChildren = ["left", "center"]; grpLine.spacing = 5;
    var lblTitle = grpLine.add("statictext", undefined, title); lblTitle.preferredSize.width = 50;
    var previewBox = grpLine.add("panel"); previewBox.preferredSize = [20, 40];
    var grpSliders = grpLine.add("group"); grpSliders.orientation = "column"; grpSliders.spacing = 0; grpSliders.margins = 0;
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

function main() {
    
    // --- LECTURE CSV (Une seule fois) ---
    var csvFile;
    if ($.os.indexOf("Windows") !== -1) csvFile = File.openDialog("Sélectionnez CSV", "*.csv;*.txt;*.*"); 
    else csvFile = File.openDialog("Sélectionnez CSV", function(f) { return true; }); 
    if (!csvFile) return;

    csvFile.open("r"); var content = csvFile.read(); csvFile.close();
    var lines = content.split(/\r\n|\r|\n/);
    if (lines.length < 1) return;
    var separator = (lines[0].indexOf(";") > -1) ? ";" : ",";
    var data = [];
    for (var i = 0; i < lines.length; i++) { if (lines[i] !== "") data.push(lines[i].split(separator)); }
    if(data.length === 0) return;
    var headers = data[0]; 

    // --- CONSTRUCTION FENETRE ---
    var win = new Window("palette", "Tableau Pro V14");
    win.orientation = "column"; win.spacing = 5; win.margins = 5;

    var grpMain = win.add("group"); 
    grpMain.orientation = "row"; grpMain.alignChildren = ["left", "top"]; grpMain.spacing = 10;

    // Colonne Params
    var colLeft = grpMain.add("group"); colLeft.orientation = "column"; colLeft.spacing = 2;
    var pGeo = colLeft.add("panel", undefined, "Géométrie"); pGeo.alignChildren = "left"; pGeo.margins = 5;
    var addSmallInput = function(p, lbl, val) {
        var g = p.add("group"); g.spacing = 2;
        g.add("statictext", undefined, lbl).preferredSize.width = 45;
        var i = g.add("edittext", undefined, val); i.preferredSize.width = 30; return i;
    };
    var inpHeadH = addSmallInput(pGeo, "H.Titre:", "40");
    var inpRowH  = addSmallInput(pGeo, "H.Ligne:", "20");
    var inpPad   = addSmallInput(pGeo, "Marge:", "15");

    var pTypo = colLeft.add("panel", undefined, "Typo"); pTypo.alignChildren = "left"; pTypo.margins = 5;
    var gF = pTypo.add("group"); gF.spacing=2;
    gF.add("statictext", undefined, "Font:");
    var inpFont = gF.add("edittext", undefined, "ArialMT"); inpFont.preferredSize.width = 60;
    var inpSize = addSmallInput(pTypo, "Taille:", "10");
    var chkBold = pTypo.add("checkbox", undefined, "Gras (Titre)");

    var pCols = colLeft.add("panel", undefined, "Colonnes"); pCols.margins = 5; pCols.maximumSize.height = 200; 
    var colScroll = pCols.add("group"); colScroll.orientation = "column"; colScroll.alignChildren = "left"; colScroll.spacing = 0;
    var checkBoxes = [];
    for (var h = 0; h < headers.length; h++) {
        var cb = colScroll.add("checkbox", undefined, headers[h]); cb.value = true; checkBoxes.push(cb);
    }

    // Colonne Couleurs
    var colRight = grpMain.add("group"); colRight.orientation = "column";
    var tPanel = colRight.add("tabbedpanel"); tPanel.preferredSize = [220, 250];
    
    var tabHead = tPanel.add("tab", undefined, "Entête"); tabHead.alignChildren = ["fill", "top"]; tabHead.margins = 5;
    var uiTxtHead = createCompactColorBlock(tabHead, "Texte", 0, 0, 0, 0);
    var uiBgHead = createCompactColorBlock(tabHead, "Fond", 0, 0, 0, 80);

    var tabCont = tPanel.add("tab", undefined, "Contenu"); tabCont.alignChildren = ["fill", "top"]; tabCont.margins = 5;
    var uiTxtCont = createCompactColorBlock(tabCont, "Texte", 0, 0, 0, 100);
    var uiBgRow1 = createCompactColorBlock(tabCont, "Ligne 1", 0, 0, 0, 0);
    var uiBgRow2 = createCompactColorBlock(tabCont, "Ligne 2", 0, 0, 0, 10);
    tPanel.selection = 0;

    // Boutons
    var gBtn = win.add("group"); gBtn.spacing = 5;
    var btnPreview = gBtn.add("button", undefined, "Preview"); btnPreview.preferredSize.width = 70;
    var btnOK = gBtn.add("button", undefined, "OK"); btnOK.preferredSize.width = 50;
    var btnCancel = gBtn.add("button", undefined, "X"); btnCancel.preferredSize.width = 30;


    // --- LA CLE DU SUCCES : BRIDGETALK ---
    var sendCommand = function(isPreview) {
        // 1. On vérifie s'il y a un doc (dans le doute, on alerte ici)
        if (app.documents.length === 0) {
            alert("Veuillez créer un document.");
            return;
        }

        // 2. On rassemble les données
        var activeIndices = [];
        for (var c = 0; c < checkBoxes.length; c++) if (checkBoxes[c].value) activeIndices.push(c);

        var params = {
            data: data,
            activeCols: activeIndices,
            isPreview: isPreview,
            geo: {
                hHead: parseFloat(inpHeadH.text)||40,
                hRow: parseFloat(inpRowH.text)||20,
                pad: parseFloat(inpPad.text)||15,
                fSize: parseFloat(inpSize.text)||10,
                fName: inpFont.text,
                isBold: chkBold.value
            },
            colors: {
                txtHead: [parseFloat(uiTxtHead.c.input.text), parseFloat(uiTxtHead.m.input.text), parseFloat(uiTxtHead.y.input.text), parseFloat(uiTxtHead.k.input.text)],
                bgHead:  [parseFloat(uiBgHead.c.input.text), parseFloat(uiBgHead.m.input.text), parseFloat(uiBgHead.y.input.text), parseFloat(uiBgHead.k.input.text)],
                txtCont: [parseFloat(uiTxtCont.c.input.text), parseFloat(uiTxtCont.m.input.text), parseFloat(uiTxtCont.y.input.text), parseFloat(uiTxtCont.k.input.text)],
                bgRow1:  [parseFloat(uiBgRow1.c.input.text), parseFloat(uiBgRow1.m.input.text), parseFloat(uiBgRow1.y.input.text), parseFloat(uiBgRow1.k.input.text)],
                bgRow2:  [parseFloat(uiBgRow2.c.input.text), parseFloat(uiBgRow2.m.input.text), parseFloat(uiBgRow2.y.input.text), parseFloat(uiBgRow2.k.input.text)]
            }
        };

        // 3. On envoie le paquet à Illustrator via BridgeTalk
        // Cela force l'exécution dans le contexte du document, pas de la fenêtre
        var bt = new BridgeTalk();
        bt.target = "illustrator";
        // On sérialise la fonction et les params pour les envoyer
        bt.body = "(" + illustratorEngine.toString() + ")(" + params.toSource() + ");";
        bt.send();
    };

    btnPreview.onClick = function() { sendCommand(true); };

    btnOK.onClick = function() {
        sendCommand(false); // Dessin final
        win.close();
    };

    btnCancel.onClick = function() {
        // On envoie une mini commande pour nettoyer si besoin
        var bt = new BridgeTalk();
        bt.target = "illustrator";
        bt.body = "try{ app.activeDocument.pageItems.getByName('Tableau_PREVIEW_TEMP').remove(); }catch(e){}";
        bt.send();
        win.close();
    };

    win.show();
}

main();