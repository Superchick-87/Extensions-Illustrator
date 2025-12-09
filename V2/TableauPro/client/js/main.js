/* client/js/main.js - VERSION V23 : SUPPRESSION STYLES INLINE (PRIORITÉ AU CSS) */

// --- 1. VARIABLES GLOBALES ---
var fullRawData = []; 
var loadedData = [];  
var headers = [];
var globalFontData = {}; 
var uncheckedColIndices = []; 

// --- VARIABLES UI ---
var selectedUiIndices = []; 
var lastClickedIndex = -1;  
var isDragging = false; 

// --- GESTION CONFLITS ---
var lastSelectionData = ""; 
var isUserInteracting = false; 
var isScriptUpdating = false;
var isFetching = false; 
var updateTimeout = null;

var interactTimer;
function userAction() {
    isUserInteracting = true;
    clearTimeout(interactTimer);
    interactTimer = setTimeout(function(){ isUserInteracting = false; }, 3000);
}

// --- 2. STYLE DYNAMIQUE ---
function injectStyles() {
    // NETTOYAGE : On ne met plus rien ici.
    // Tous les styles (.col-item, .edit-icon, etc.) sont désormais gérés
    // proprement dans le fichier css/style.css pour éviter les conflits.
}

// --- 3. UTILITAIRES ---
function setupShiftIncrement(input) {
    if(!input || input.type !== 'number') return;
    if (input.dataset.shiftAttached === "true") return;
    input.dataset.shiftAttached = "true";

    input.addEventListener('keydown', function(e) {
        if (e.shiftKey && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
            e.preventDefault();
            e.stopImmediatePropagation();
            var currentVal = parseFloat(this.value) || 0;
            var step = 10;
            if (e.key === "ArrowDown") step = -10;
            var newVal = currentVal + step;
            if(this.min !== "" && newVal < parseFloat(this.min)) newVal = parseFloat(this.min);
            if(this.max !== "" && newVal > parseFloat(this.max)) newVal = parseFloat(this.max);
            this.value = newVal;
            this.dispatchEvent(new Event('input'));
            this.dispatchEvent(new Event('change'));
        }
    });
}

function runJSX(script, callback) {
    if (window.__adobe_cep__ && window.__adobe_cep__.evalScript) {
        window.__adobe_cep__.evalScript(script, callback);
    } else if (window.cep && window.cep.process && window.cep.process.evalScript) {
        window.cep.process.evalScript(script, callback);
    } 
}

function toHex(str) {
    try {
        var utf8 = unescape(encodeURIComponent(str));
        var r = ''; for (var i=0; i<utf8.length; i++) { var h = utf8.charCodeAt(i).toString(16); r+=("0"+h).slice(-2); }
        return r.toUpperCase();
    } catch (e) { return ""; }
}

function cmykToHex(c, m, y, k) {
    var r = 255 * (1 - c / 100) * (1 - k / 100);
    var g = 255 * (1 - m / 100) * (1 - k / 100);
    var b = 255 * (1 - y / 100) * (1 - k / 100);
    function compToHex(c) {
        var hex = Math.round(c).toString(16);
        return hex.length == 1 ? "0" + hex : hex;
    }
    return "#" + compToHex(r) + compToHex(g) + compToHex(b);
}

function hexToCmyk(hex) {
    var r = parseInt(hex.substring(1, 3), 16) / 255;
    var g = parseInt(hex.substring(3, 5), 16) / 255;
    var b = parseInt(hex.substring(5, 7), 16) / 255;
    var k = 1 - Math.max(r, g, b);
    var c = (1 - r - k) / (1 - k) || 0;
    var m = (1 - g - k) / (1 - k) || 0;
    var y = (1 - b - k) / (1 - k) || 0;
    return [Math.round(c * 100), Math.round(m * 100), Math.round(y * 100), Math.round(k * 100)];
}

// --- 4. LOGIQUE UI SPÉCIFIQUE ---

function checkLegendState() {
    var legText = document.getElementById('legendText').value;
    var hasLegend = legText && legText.trim() !== "";
    var container = document.getElementById('legRadiusContainer');
    var syncCb = document.getElementById('syncRadius');
    if (!container || !syncCb) return;
    var inputs = container.querySelectorAll('input');
    if (hasLegend) {
        container.style.opacity = "1"; container.style.pointerEvents = "auto";
        inputs.forEach(i => i.disabled = false); syncCb.disabled = false;
    } else {
        container.style.opacity = "0.3"; container.style.pointerEvents = "none";
        inputs.forEach(i => i.disabled = true); syncCb.checked = false; syncCb.disabled = true;
    }
}

function syncRadiusLogic(source) {
    var tTL = document.getElementById('radTL').value; var tTR = document.getElementById('radTR').value;
    var tBL = document.getElementById('radBL').value; var tBR = document.getElementById('radBR').value;
    var tGlob = document.getElementById('radGlobal');
    
    if (source === 'radGlobal') { ['radTL', 'radTR', 'radBL', 'radBR'].forEach(id => document.getElementById(id).value = tGlob.value); } 
    else { if (tTL === tTR && tTR === tBL && tBL === tBR) tGlob.value = tTL; else tGlob.value = ""; }

    var lTL = document.getElementById('legRadTL').value; var lTR = document.getElementById('legRadTR').value;
    var lBL = document.getElementById('legRadBL').value; var lBR = document.getElementById('legRadBR').value;
    var lGlob = document.getElementById('legRadGlobal');

    if (source === 'legRadGlobal') { ['legRadTL', 'legRadTR', 'legRadBL', 'legRadBR'].forEach(id => document.getElementById(id).value = lGlob.value); } 
    else { if (lTL === lTR && lTR === lBL && lBL === lBR) lGlob.value = lTL; else lGlob.value = ""; }

    var isSynced = document.getElementById('syncRadius').checked;
    if (isSynced) {
        if (source.startsWith('rad')) {
            document.getElementById('legRadTL').value = document.getElementById('radTL').value;
            document.getElementById('legRadTR').value = document.getElementById('radTR').value;
            document.getElementById('legRadBL').value = document.getElementById('radBL').value;
            document.getElementById('legRadBR').value = document.getElementById('radBR').value;
            document.getElementById('legRadGlobal').value = document.getElementById('radGlobal').value;
        }
        if (source === 'syncRadius') {
            document.getElementById('legRadTL').value = document.getElementById('radTL').value;
            document.getElementById('legRadTR').value = document.getElementById('radTR').value;
            document.getElementById('legRadBL').value = document.getElementById('radBL').value;
            document.getElementById('legRadBR').value = document.getElementById('radBR').value;
            document.getElementById('legRadGlobal').value = document.getElementById('radGlobal').value;
        }
    }
}

function updateMasterCheck() {
    var bAll = document.getElementById('bAll');
    if (!bAll) return;
    var ids = ['bVert', 'bHorz', 'bHead', 'bOuter'];
    var allChecked = true;
    ids.forEach(function(id) {
        var el = document.getElementById(id);
        if (el && !el.checked) allChecked = false;
    });
    bAll.checked = allChecked;
}

// --- 5. GÉNÉRATION UI COLONNES (AVEC RENOMMAGE) ---
// --- 5. GÉNÉRATION UI COLONNES (MULTI-DRAG & DROP) ---
function generateColUI(preserveState) {
    var lst = document.getElementById('colList'); if(!lst) return; 
    
    var savedAligns = {};
    if (preserveState && lst.children.length > 0) {
        lst.querySelectorAll('.col-item').forEach(function(item) {
            savedAligns[parseInt(item.getAttribute('data-idx'))] = item.querySelector('.align-btn.selected').getAttribute('data-al');
        });
    }

    lst.innerHTML = "";
    
    // Fonction helper pour vérifier si un index est sélectionné
    function isUiSelected(idx) { return selectedUiIndices.includes(idx); }

    headers.forEach(function(h, idx){
        var safe = String(h).replace(/^"|"$/g,'').trim() || ("Col "+(idx+1));
        var isChecked = !uncheckedColIndices.includes(idx);
        var alignState = savedAligns[idx] || 'left';

        var el = document.createElement('div'); el.className = 'col-item';
        if (isUiSelected(idx)) el.classList.add('ui-selected'); 
        
        el.setAttribute('draggable','true'); 
        el.setAttribute('data-idx', idx);
        
        var checkAttr = isChecked ? 'checked' : '';
        var clL = (alignState === 'left') ? 'selected' : '';
        var clC = (alignState === 'center') ? 'selected' : '';
        var clR = (alignState === 'right') ? 'selected' : '';

        // HTML
        el.innerHTML = `
            <div class="drag-handle">☰</div>
            <input type="checkbox" ${checkAttr}>
            <span class="col-name" title="${safe}">${safe}</span>
            <span class="edit-icon" title="Renommer">✎</span>
            <div class="align-grp">
                <div class="align-btn ${clL}" data-al="left">L</div>
                <div class="align-btn ${clC}" data-al="center">C</div>
                <div class="align-btn ${clR}" data-al="right">R</div>
            </div>`;
        
        // --- RENOMMAGE ---
        function startRename() {
            var nameSpan = el.querySelector('.col-name');
            var currentName = headers[idx]; 
            var input = document.createElement('input');
            input.type = "text";
            input.value = currentName.replace(/^"|"$/g,'');
            input.className = "col-name-input";
            nameSpan.replaceWith(input);
            input.focus();
            function validate() {
                var newVal = input.value;
                headers[idx] = newVal;
                if(loadedData[0]) loadedData[0][idx] = newVal;
                generateColUI(true);
                userAction();
            }
            input.onblur = validate;
            input.onkeydown = function(e) { if(e.key === "Enter") { validate(); } e.stopPropagation(); };
            input.onclick = function(e) { e.stopPropagation(); };
        }
        var nameEl = el.querySelector('.col-name');
        var editIcon = el.querySelector('.edit-icon');
        nameEl.ondblclick = function(e) { e.stopPropagation(); startRename(); };
        editIcon.onclick = function(e) { e.stopPropagation(); startRename(); };

        // --- CLIC SELECTION ---
        el.onclick = function(e) {
            if (isDragging) return; 
            if (e.target.tagName === 'INPUT' || e.target.classList.contains('align-btn') || e.target.classList.contains('edit-icon') || e.target.classList.contains('drag-handle') || e.target.classList.contains('col-name-input')) return;

            var currentIdx = parseInt(el.getAttribute('data-idx'));
            if (e.shiftKey && lastClickedIndex !== -1) {
                var start = Math.min(lastClickedIndex, currentIdx);
                var end = Math.max(lastClickedIndex, currentIdx);
                selectedUiIndices = [];
                for(var i=start; i<=end; i++) { selectedUiIndices.push(i); }
            } else if (e.ctrlKey || e.metaKey) {
                if (selectedUiIndices.includes(currentIdx)) {
                    selectedUiIndices = selectedUiIndices.filter(function(id){ return id !== currentIdx; });
                } else { selectedUiIndices.push(currentIdx); }
                lastClickedIndex = currentIdx;
            } else {
                selectedUiIndices = [currentIdx];
                lastClickedIndex = currentIdx;
            }
            generateColUI(true); 
        };

        // EVENT LISTENERS INTERNES
        var cb = el.querySelector('input[type="checkbox"]');
        cb.onclick = function(e) {
            e.stopPropagation();
            var targetState = this.checked;
            var currentIdx = parseInt(el.getAttribute('data-idx'));
            var itemsToProcess = isUiSelected(currentIdx) ? selectedUiIndices : [currentIdx];
            itemsToProcess.forEach(function(targetIdx) {
                var domItem = lst.querySelector(`.col-item[data-idx="${targetIdx}"]`);
                if (domItem) {
                    domItem.querySelector('input[type="checkbox"]').checked = targetState; 
                    if (targetState) { uncheckedColIndices = uncheckedColIndices.filter(function(id) { return id !== targetIdx; }); } 
                    else { if (!uncheckedColIndices.includes(targetIdx)) uncheckedColIndices.push(targetIdx); }
                }
            });
            userAction();
        };

        el.querySelectorAll('.align-btn').forEach(function(b){ 
            b.onclick=function(e){ 
                e.stopPropagation();
                var newAlign = this.getAttribute('data-al');
                var currentIdx = parseInt(el.getAttribute('data-idx'));
                var itemsToProcess = isUiSelected(currentIdx) ? selectedUiIndices : [currentIdx];
                itemsToProcess.forEach(function(targetIdx) {
                    var domItem = lst.querySelector(`.col-item[data-idx="${targetIdx}"]`);
                    if (domItem) {
                        domItem.querySelectorAll('.align-btn').forEach(function(x){ x.classList.remove('selected'); });
                        domItem.querySelector(`.align-btn[data-al="${newAlign}"]`).classList.add('selected');
                    }
                });
                userAction(); 
            }
        });
        
        // --- DRAG START (MULTI) ---
        el.addEventListener('dragstart', function(e){ 
            isDragging = true;
            
            var currentIdx = parseInt(el.getAttribute('data-idx'));
            
            // Si on drag un élément qui n'est PAS sélectionné, on le sélectionne seul
            if (!selectedUiIndices.includes(currentIdx)) {
                selectedUiIndices = [currentIdx];
                // Mise à jour visuelle rapide
                lst.querySelectorAll('.col-item').forEach(it => it.classList.remove('ui-selected'));
                el.classList.add('ui-selected');
            }

            // On ajoute la classe .dragging à TOUS les éléments sélectionnés
            lst.querySelectorAll('.col-item.ui-selected').forEach(function(item) {
                item.classList.add('dragging');
            });

            if(e.dataTransfer) { 
                e.dataTransfer.effectAllowed = 'move'; 
                e.dataTransfer.setData('text/html', null); 
                // Image fantôme (optionnel, prend le premier élément)
                e.dataTransfer.setDragImage(el, 0, 0);
            }
            userAction(); 
        });
        
        // --- DRAG END ---
        el.addEventListener('dragend', function(){ 
            lst.querySelectorAll('.dragging').forEach(it => it.classList.remove('dragging'));
            reorderColumnsData(); 
            setTimeout(function(){ isDragging = false; }, 50); 
        });

        lst.appendChild(el);
    });
    
    // --- DRAG OVER (MULTI) ---
    lst.ondragover = function(e){
        e.preventDefault(); 
        
        // On récupère TOUS les éléments en cours de déplacement
        var draggingItems = document.querySelectorAll('.dragging');
        if(draggingItems.length === 0) return;

        // On cherche l'élément cible sous la souris (qui n'est pas en train d'être déplacé)
        var siblings = [...lst.querySelectorAll('.col-item:not(.dragging)')];
        var nextSibling = siblings.find(s => e.clientY <= s.getBoundingClientRect().top + s.offsetHeight/2);
        
        // On déplace le GROUPE entier
        draggingItems.forEach(function(dragItem) {
            if(nextSibling) lst.insertBefore(dragItem, nextSibling);
            else lst.appendChild(dragItem);
        });
    };
}

// --- FONCTION REORGANIATION DONNEES ---
function reorderColumnsData() {
    var lst = document.getElementById('colList');
    var newOrderIndices = [];
    lst.querySelectorAll('.col-item').forEach(function(item) {
        newOrderIndices.push(parseInt(item.getAttribute('data-idx')));
    });

    var newHeaders = newOrderIndices.map(i => headers[i]);
    
    var newLoadedData = [];
    newLoadedData.push(newHeaders);
    
    for (var r = 1; r < loadedData.length; r++) {
        var oldRow = loadedData[r];
        var newRow = newOrderIndices.map(i => oldRow[i]);
        newLoadedData.push(newRow);
    }
    
    headers = newHeaders;
    loadedData = newLoadedData;
    
    uncheckedColIndices = [];
    lst.querySelectorAll('.col-item').forEach(function(item, newIndex) {
        if (!item.querySelector('input[type="checkbox"]').checked) {
            uncheckedColIndices.push(newIndex);
        }
        item.setAttribute('data-idx', newIndex);
    });

    selectedUiIndices = [];
    generateColUI(true);
}

// --- 6. PROCESS DATA ---
function processDataRange() {
    if (fullRawData.length === 0) return;

    var inputStart = document.getElementById('rowStart');
    var inputEnd = document.getElementById('rowEnd');
    var cbNbLines = document.getElementById('cbNbLines');
    var valNbLines = document.getElementById('valNbLines');
    var warningEl = document.getElementById('rangeWarning');
    var btnApply = document.getElementById('btnApply');

    var start = parseInt(inputStart.value) || 0;
    var end = parseInt(inputEnd.value) || 0;
    var nbLines = parseInt(valNbLines.value) || 0;
    var useNbLines = cbNbLines ? cbNbLines.checked : false;

    if (!useNbLines && end > 0 && start > end) {
        if(warningEl) warningEl.style.display = 'block';
        if(btnApply) btnApply.disabled = true;
        return; 
    } else {
        if(warningEl) warningEl.style.display = 'none';
        if(btnApply) btnApply.disabled = false;
    }

    var csvHeader = fullRawData[0];
    var csvRows = fullRawData.slice(1);
    var slicedRows = csvRows;

    if (start > 0) {
        var startIndex = Math.max(0, start - 1);
        slicedRows = slicedRows.slice(startIndex);
    }

    if (useNbLines) {
        if (nbLines > 0) slicedRows = slicedRows.slice(0, nbLines);
    } else {
        if (end > 0) {
            var lengthToKeep = end;
            if (start > 0) lengthToKeep = (end - start) + 1;
            if (lengthToKeep < 0) lengthToKeep = 0;
            slicedRows = slicedRows.slice(0, lengthToKeep);
        }
    }

    loadedData = [csvHeader].concat(slicedRows);
    headers = csvHeader;
    generateColUI(true);
}


// --- 7. UI COULEURS ---
function createCMYKControl(id, t, v) {
    var d = document.getElementById(id); if(!d) return;
    var h = `<div class="cmyk-title">${t}</div><div class="cmyk-wrapper"><div class="cmyk-body"><div style="display:flex; flex-direction:column; gap:2px;">`; 
    var l=['C','M','J','N'];
    for(var i=0;i<4;i++){
        h+=`<div class="cmyk-row"><span class="cmyk-label" style="width:15px; font-size:10px; color:#aaa;">${l[i]}</span><input type="range" min="0" max="100" value="${v[i]}" class="sl-${i}"><div class="number-wrapper"><input type="number" min="0" max="100" value="${v[i]}" class="inp-${i}"><span class="unit-symbol">%</span></div></div>`;
    }
    h += `</div><div><div class="swatch-visual" style="width:25px; height:25px; border:1px solid #555; background-color:#000; cursor:pointer;" title="Changer la couleur"></div><input type="color" class="color-hidden" style="display:none;" value="#000000"></div></div></div>`;
    d.innerHTML = h;
    var swatch = d.querySelector('.swatch-visual');
    var hiddenPicker = d.querySelector('.color-hidden');
    swatch.onclick = function() { hiddenPicker.click(); };
    function updatePreview() {
        var vals = getCMYK(id);
        var hex = cmykToHex(vals[0], vals[1], vals[2], vals[3]);
        hiddenPicker.value = hex;
        swatch.style.backgroundColor = hex;
    }
    function updateFromPicker() {
        var hex = hiddenPicker.value;
        swatch.style.backgroundColor = hex;
        var cmyk = hexToCmyk(hex);
        setCMYK(id, cmyk);
        userAction(); 
    }
    hiddenPicker.oninput = function() { updateFromPicker(); userAction(); };
    hiddenPicker.onchange = function() { updateFromPicker(); userAction(); };
    var ranges = d.querySelectorAll('input[type="range"]');
    var numbers = d.querySelectorAll('input[type="number"]');
    ranges.forEach(function(rng){
        rng.onmousedown = function() { userAction(); };
        rng.oninput = function() { this.nextElementSibling.querySelector('input').value = this.value; updatePreview(); userAction(); };
    });
    numbers.forEach(function(num){
        setupShiftIncrement(num);
        num.onchange = function() { var val = parseInt(this.value); if(isNaN(val)) val = 0; if(val > 100) val = 100; if(val < 0) val = 0; this.value = val; this.parentElement.previousElementSibling.value = val; updatePreview(); userAction(); };
        num.oninput = function() { if (this.value === "") { updatePreview(); return; } var val = parseInt(this.value); if(val > 100) { val = 100; this.value = 100; } this.parentElement.previousElementSibling.value = val; updatePreview(); userAction(); };
    });
    updatePreview();
}

function setCMYK(id, values) {
    var d = document.getElementById(id); if (!d) return;
    for(var i=0; i<4; i++) {
        var val = values[i] || 0;
        d.querySelector(`.sl-${i}`).value = val;
        d.querySelector(`.inp-${i}`).value = val;
    }
    var hex = cmykToHex(values[0]||0, values[1]||0, values[2]||0, values[3]||0);
    var pk = d.querySelector('.color-hidden');
    var sw = d.querySelector('.swatch-visual');
    if(pk) pk.value = hex;
    if(sw) sw.style.backgroundColor = hex;
}

function getCMYK(id) {
    var d=document.getElementById(id); var v=[];
    for(var i=0;i<4;i++) { var val = parseInt(d.querySelector(`.inp-${i}`).value); if(isNaN(val)) val = 0; v.push(val); }
    return v;
}


// --- 8. TABS, FONTS & WATCHER ---

function initTabs() {
    var mainTabs = document.querySelectorAll('.tab-btn');
    mainTabs.forEach(function(btn) {
        btn.onclick = function() {
            mainTabs.forEach(function(b) { b.classList.remove('active'); });
            document.querySelectorAll('.tab-content').forEach(function(c) { c.classList.remove('active'); });
            this.classList.add('active');
            var targetId = this.getAttribute('data-tab');
            var targetContent = document.getElementById(targetId);
            if(targetContent) targetContent.classList.add('active');
        };
    });

    var subTabs = document.querySelectorAll('.sub-tab-btn');
    subTabs.forEach(function(btn) {
        btn.onclick = function() {
            var parentGroup = this.parentElement;
            parentGroup.querySelectorAll('.sub-tab-btn').forEach(function(b) { b.classList.remove('active'); });
            this.classList.add('active');
            var targetId = this.getAttribute('data-subtab');
            var targetContent = document.getElementById(targetId);
            if(targetContent) {
                var contentParent = targetContent.parentElement;
                var siblings = contentParent.children;
                for (var i = 0; i < siblings.length; i++) {
                    if (siblings[i].classList.contains('sub-tab-content')) {
                        siblings[i].classList.remove('active');
                    }
                }
                targetContent.classList.add('active');
            }
        };
    });
}

function loadFonts() {
    var originalSelect = document.getElementById('fName'); 
    var container = document.getElementById('font-ui-container');

    if (!container && originalSelect) {
        originalSelect.style.display = 'none'; 
        container = document.createElement('div');
        container.id = 'font-ui-container';
        container.style.display = 'flex';
        container.style.flexDirection = 'column';
        container.style.gap = '5px';

        var selFull = document.createElement('select');
        selFull.id = 'ui-font-full';
        selFull.style.width = '240px';

        var filterInput = document.getElementById("fontFilter");
        if(filterInput) {
            filterInput.oninput = function() {
                var q = this.value.toLowerCase();
                for (var i = 0; i < selFull.options.length; i++) {
                    var opt = selFull.options[i];
                    opt.style.display = opt.text.toLowerCase().includes(q) ? "block" : "none";
                }
            };
        }

        originalSelect.parentNode.insertBefore(container, originalSelect);
        container.appendChild(selFull);
        selFull.onchange = function() {
            if (isScriptUpdating) return;
            userAction();
            originalSelect.value = this.value; 
        };
    }

    var selFull = document.getElementById('ui-font-full');
    if (selFull) selFull.innerHTML = "<option>Chargement...</option>";
    if (originalSelect) originalSelect.innerHTML = "<option>Chargement...</option>";

    runJSX('getSystemFonts()', function(res) {
        if (!res) return;
        var fonts = res.split("|").filter(function(f) { return f && f.trim().length > 0; }).sort();
        if(selFull) selFull.innerHTML = "";
        if(originalSelect) originalSelect.innerHTML = "";
        fonts.forEach(function(f) {
            var opt = document.createElement('option');
            opt.value = f; opt.innerText = f;
            if(selFull) selFull.appendChild(opt);
            var optHidden = document.createElement('option');
            optHidden.value = f; optHidden.innerText = f;
            if(originalSelect) originalSelect.appendChild(optHidden);
        });
        if (fonts.length > 0) {
            if(selFull) selFull.value = fonts[0];
            if(originalSelect) originalSelect.value = fonts[0];
        }
        var savedFont = originalSelect.getAttribute('data-saved');
        if(savedFont && fonts.includes(savedFont)) {
            selFull.value = savedFont;
            originalSelect.value = savedFont;
        }
    });
}

function checkForUpdates() {
    if (isFetching) return;
    isFetching = true;
    runJSX('getSelectionData()', function(res) {
        isFetching = false;
        if (!res || res === "" || res === lastSelectionData) {
            if(res === "") lastSelectionData = "";
            return;
        }
        lastSelectionData = res;
        updateUIFromData(res);
    });
}

function startSelectionWatcher() {
    try {
        var csInterface = new CSInterface();
        csInterface.addEventListener("documentAfterSelectionChanged", function() {
            clearTimeout(updateTimeout);
            updateTimeout = setTimeout(checkForUpdates, 10);
        });
    } catch(e) {}
    setInterval(checkForUpdates, 200);
}

function updateUIFromData(hexData) {
    try {
        isScriptUpdating = true;
        var tab = document.querySelector('[data-tab="tab-style"]');
        if (tab) tab.click();

        var hexArr = [];
        for (var i = 0; i < hexData.length; i += 2) {
            hexArr.push(String.fromCharCode(parseInt(hexData.substr(i, 2), 16)));
        }
        var p = JSON.parse(decodeURIComponent(escape(hexArr.join(''))));

        if (p.data && p.data.length > 0) {
            var isNewData = JSON.stringify(p.data) !== JSON.stringify(loadedData);
            loadedData = p.data;
            headers = loadedData[0];
            uncheckedColIndices = [];
            if (p.activeCols) {
                headers.forEach(function(_, idx) {
                    if (!p.activeCols.includes(idx))
                        uncheckedColIndices.push(idx);
                });
            }
            generateColUI(false);
            if (p.colAligns && p.activeCols) {
                p.activeCols.forEach(function(colIdx, i) {
                    var item = document.querySelector(`.col-item[data-idx="${colIdx}"]`);
                    if (!item) return;
                    item.querySelectorAll('.align-btn').forEach(btn => btn.classList.remove('selected'));
                    var b = item.querySelector(`.align-btn[data-al="${p.colAligns[i]}"]`);
                    if (b) b.classList.add('selected');
                });
            }
        }

        if (p.csvOptions) {
            if(document.getElementById('rowStart')) document.getElementById('rowStart').value = p.csvOptions.start || 0;
            if(document.getElementById('rowEnd')) document.getElementById('rowEnd').value = p.csvOptions.end || 0;
            if(document.getElementById('valNbLines')) document.getElementById('valNbLines').value = p.csvOptions.nbLines || 10;
            
            var cbNbLines = document.getElementById('cbNbLines');
            if(cbNbLines) {
                cbNbLines.checked = p.csvOptions.useNbLines || false;
                cbNbLines.dispatchEvent(new Event('change'));
            }
            var cbShowTitle = document.getElementById('cbShowTitle');
            if(cbShowTitle) {
                cbShowTitle.checked = (p.csvOptions.showTitle !== undefined) ? p.csvOptions.showTitle : true;
            }
            if (p.csvOptions.legend !== undefined) {
                var legArea = document.getElementById('legendText');
                if(legArea) legArea.value = p.csvOptions.legend;
            }
        }

        if (p.geo) {
            if(document.getElementById('fSize')) document.getElementById('fSize').value = p.geo.fSize;
            if(document.getElementById('fLeading')) document.getElementById('fLeading').value = p.geo.leading;
            if(document.getElementById('hHead')) document.getElementById('hHead').value = p.geo.hHead;
            if(document.getElementById('hRow')) document.getElementById('hRow').value = p.geo.hRow;
            if(document.getElementById('pad')) document.getElementById('pad').value = p.geo.pad;
            if(document.getElementById('isBold')) document.getElementById('isBold').checked = p.geo.isBold;
            if(document.getElementById('wrapHead')) document.getElementById('wrapHead').checked = p.geo.wrapHead;
            if (p.geo && p.geo.fName) {
                var fullFont = p.geo.fName;
                var selFull = document.getElementById('ui-font-full');
                var originalSelect = document.getElementById('fName');
                if (selFull) selFull.value = fullFont;
                if (originalSelect) originalSelect.value = fullFont;
            }
            // RADIUS UPDATE UI
            if (p.geo.radius) {
                if(document.getElementById('radTL')) document.getElementById('radTL').value = p.geo.radius.tl || 0;
                if(document.getElementById('radTR')) document.getElementById('radTR').value = p.geo.radius.tr || 0;
                if(document.getElementById('radBL')) document.getElementById('radBL').value = p.geo.radius.bl || 0;
                if(document.getElementById('radBR')) document.getElementById('radBR').value = p.geo.radius.br || 0;
            }
            if (p.geo.legRadius) {
                if(document.getElementById('legRadTL')) document.getElementById('legRadTL').value = p.geo.legRadius.tl || 0;
                if(document.getElementById('legRadTR')) document.getElementById('legRadTR').value = p.geo.legRadius.tr || 0;
                if(document.getElementById('legRadBL')) document.getElementById('legRadBL').value = p.geo.legRadius.bl || 0;
                if(document.getElementById('legRadBR')) document.getElementById('legRadBR').value = p.geo.legRadius.br || 0;
            }
        }
        if (p.colors) {
            setCMYK('ui-cTxtHead', p.colors.txtHead); 
            setCMYK('ui-cBgHead', p.colors.bgHead);
            setCMYK('ui-cTxtCont', p.colors.txtCont); 
            setCMYK('ui-cBgRow1', p.colors.bgRow1);
            setCMYK('ui-cBgRow2', p.colors.bgRow2);
            if (p.colors.bgLeg) setCMYK('ui-cBgLeg', p.colors.bgLeg);
            if (p.colors.txtLeg) setCMYK('ui-cTxtLeg', p.colors.txtLeg);
            if (p.colors.strokeHead) setCMYK('ui-cStrokeHead', p.colors.strokeHead);
        }
        if (p.borders) {
            if(document.getElementById('bVert')) document.getElementById('bVert').checked = p.borders.vert;
            if(document.getElementById('bHorz')) document.getElementById('bHorz').checked = p.borders.horz;
            if(document.getElementById('bHead')) document.getElementById('bHead').checked = p.borders.head;
            if(document.getElementById('bOuter')) document.getElementById('bOuter').checked = p.borders.outer;
            if(document.getElementById('bWeight')) document.getElementById('bWeight').value = p.borders.weight;
        }
        
        updateMasterCheck();
        syncRadiusLogic('');
        checkLegendState();

        isScriptUpdating = false; 
    } catch(e) { isScriptUpdating = false; }
}


// --- 10. INITIALISATION ---
function init() {
    injectStyles();

    createCMYKControl('ui-cBgHead', 'Fond', [0,0,0,80]);
    createCMYKControl('ui-cTxtHead', 'Texte', [0,0,0,0]);
    createCMYKControl('ui-cBgRow1', 'Fond ligne 1', [0,0,0,0]);
    createCMYKControl('ui-cBgRow2', 'Fond ligne 2', [0,0,0,10]);
    createCMYKControl('ui-cTxtCont', 'Texte', [0,0,0,100]);
    createCMYKControl('ui-cBgLeg', 'Fond Légende', [0,0,0,0]); 
    createCMYKControl('ui-cTxtLeg', 'Texte Légende', [0,0,0,100]); 
    createCMYKControl('ui-cStrokeHead', 'Filets Titre', [0,0,0,0]); 

    // BORDURES
    var bAll = document.getElementById('bAll');
    var borderIds = ['bVert', 'bHorz', 'bHead', 'bOuter'];
    if (bAll) {
        bAll.onchange = function() {
            var st = this.checked;
            borderIds.forEach(function(id){ var el = document.getElementById(id); if(el) el.checked = st; });
            userAction();
        };
    }
    borderIds.forEach(function(id){
        var el = document.getElementById(id);
        if(el) { el.addEventListener('change', function() { updateMasterCheck(); userAction(); }); }
    });

    // RADIUS
    var radInputs = [
        'radTL', 'radTR', 'radBL', 'radBR', 'radGlobal',
        'legRadTL', 'legRadTR', 'legRadBL', 'legRadBR', 'legRadGlobal'
    ];
    radInputs.forEach(function(id) {
        var el = document.getElementById(id);
        if(el) {
            el.addEventListener('input', function() {
                syncRadiusLogic(id);
                userAction();
            });
        }
    });
    
    var elSync = document.getElementById('syncRadius');
    if(elSync) {
        elSync.addEventListener('change', function() {
            syncRadiusLogic('syncRadius');
            userAction();
        });
    }
    
    var legArea = document.getElementById('legendText');
    if(legArea) {
        legArea.addEventListener('input', function() {
            checkLegendState();
            userAction();
        });
    }

    initTabs();
    
    loadFonts();
    
    var allNumberInputs = document.querySelectorAll('input[type="number"]');
    allNumberInputs.forEach(setupShiftIncrement);

    var btnImport = document.getElementById('btnImport');
    var fInput = document.getElementById('fileInput');
    
    if(btnImport && fInput) {
        btnImport.onclick = function() { fInput.click(); };
        fInput.addEventListener('change', function(e){
            var f = e.target.files[0]; if(!f) return;
            var nameDisplay = document.getElementById('fileNameDisplay');
            if(nameDisplay) nameDisplay.innerText = f.name;
            if(typeof Papa==='undefined'){alert("PapaParse manquant");return;}
            Papa.parse(f, { skipEmptyLines: true, complete: function(r){
                if(r.errors.length>0) { alert("Erreur CSV"); return; }
                fullRawData = r.data; 
                uncheckedColIndices = []; 
                processDataRange();
                var tabCols = document.querySelector('[data-tab="tab-cols"]');
                if(tabCols) tabCols.click();
            }});
        });
    }

    var inputStart = document.getElementById('rowStart');
    var inputEnd = document.getElementById('rowEnd');
    var cbNbLines = document.getElementById('cbNbLines');
    var valNbLines = document.getElementById('valNbLines');
    var cbShowTitle = document.getElementById('cbShowTitle');
    var wrapRowEnd = document.getElementById('wrapRowEnd');
    var wrapNbLines = document.getElementById('wrapNbLines');
    
    function onRangeChange() { userAction(); processDataRange(); }

    if(cbNbLines) {
        cbNbLines.onchange = function() {
            if(this.checked) {
                if(wrapRowEnd) wrapRowEnd.classList.add('input-disabled');   
                if(wrapNbLines) wrapNbLines.classList.remove('input-disabled'); 
                if(inputEnd) inputEnd.disabled = true;
                if(valNbLines) valNbLines.disabled = false;
            } else {
                if(wrapRowEnd) wrapRowEnd.classList.remove('input-disabled'); 
                if(wrapNbLines) wrapNbLines.classList.add('input-disabled');   
                if(inputEnd) inputEnd.disabled = false;
                if(valNbLines) valNbLines.disabled = true;
            }
            onRangeChange();
        };
    }

    if(inputStart) inputStart.oninput = onRangeChange;
    if(inputEnd) inputEnd.oninput = onRangeChange;
    if(valNbLines) valNbLines.oninput = onRangeChange;
    if(cbShowTitle) cbShowTitle.onchange = function() { userAction(); };

    var btnApply = document.getElementById('btnApply');
    if(btnApply) {
        btnApply.onclick = function() {
            userAction(); 
            if(loadedData.length===0){ alert("Chargez un CSV."); return; }
            var items = document.querySelectorAll('.col-item');
            var act=[]; var al=[];
            items.forEach(function(it){
                if(it.querySelector('input[type="checkbox"]').checked) {
                    act.push(parseInt(it.getAttribute('data-idx')));
                    al.push(it.querySelector('.align-btn.selected').getAttribute('data-al'));
                }
            });
            if(act.length===0){ alert("Cochez une colonne."); return; }

            var csvOpts = {
                start: parseInt(document.getElementById('rowStart').value) || 0,
                end: parseInt(document.getElementById('rowEnd').value) || 0,
                nbLines: parseInt(document.getElementById('valNbLines').value) || 0,
                useNbLines: document.getElementById('cbNbLines').checked,
                showTitle: document.getElementById('cbShowTitle').checked,
                legend: document.getElementById('legendText').value 
            };

            var p = {
                mode: "create", 
                data: loadedData,
                activeCols: act, colAligns: al,
                csvOptions: csvOpts,
                geo: {
                    fName: (document.getElementById('ui-font-full') ? document.getElementById('ui-font-full').value : document.getElementById('fName').value),
                    fSize: parseFloat(document.getElementById('fSize').value)||10,
                    leading: parseFloat(document.getElementById('fLeading').value)||12,
                    isBold: document.getElementById('isBold') ? document.getElementById('isBold').checked : false,
                    hHead: parseFloat(document.getElementById('hHead').value)||15,
                    hRow: parseFloat(document.getElementById('hRow').value)||10,
                    pad: parseFloat(document.getElementById('pad').value)||4,
                    wrapHead: document.getElementById('wrapHead') ? document.getElementById('wrapHead').checked : false,
                    maxRows: 0, totalW: 0,
                    radius: {
                        tl: parseFloat(document.getElementById('radTL').value)||0,
                        tr: parseFloat(document.getElementById('radTR').value)||0,
                        bl: parseFloat(document.getElementById('radBL').value)||0,
                        br: parseFloat(document.getElementById('radBR').value)||0
                    },
                    legRadius: {
                        tl: parseFloat(document.getElementById('legRadTL').value)||0,
                        tr: parseFloat(document.getElementById('legRadTR').value)||0,
                        bl: parseFloat(document.getElementById('legRadBL').value)||0,
                        br: parseFloat(document.getElementById('legRadBR').value)||0
                    }
                },
                colors: { 
                    txtHead: getCMYK('ui-cTxtHead'), 
                    bgHead: getCMYK('ui-cBgHead'), 
                    txtCont: getCMYK('ui-cTxtCont'), 
                    bgRow1: getCMYK('ui-cBgRow1'), 
                    bgRow2: getCMYK('ui-cBgRow2'),
                    bgLeg: getCMYK('ui-cBgLeg'),
                    txtLeg: getCMYK('ui-cTxtLeg'),
                    strokeHead: getCMYK('ui-cStrokeHead')
                },
                borders: { 
                    vert: document.getElementById('bVert') ? document.getElementById('bVert').checked : true, 
                    horz: document.getElementById('bHorz') ? document.getElementById('bHorz').checked : true, 
                    head: document.getElementById('bHead') ? document.getElementById('bHead').checked : true, 
                    outer: document.getElementById('bOuter') ? document.getElementById('bOuter').checked : true, 
                    weight: parseFloat(document.getElementById('bWeight').value)||0.5 
                }
            };

            if (!csvOpts.showTitle) {
                p.geo.hHead = 0;        
                p.borders.head = false; 
            }

            runJSX('creerTableau("' + toHex(JSON.stringify(p)) + '")', function(){});
        };
    }

    startSelectionWatcher();
}

window.onload = init;