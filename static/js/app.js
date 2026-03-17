/**
 * app.js  (v2 — Google Fonts + Bold/Italic + Custom Font Upload)
 * ---------------------------------------------------------------
 * Frontend logic for the Certificate Generator.
 */

// ============================================================
// A curated list of popular Google Fonts (300+ fonts)
// Loaded via the Google Fonts CSS API for Fabric.js canvas,
// and the font name is passed to the backend for Pillow rendering.
// ============================================================
const GOOGLE_FONTS = [
    "Roboto", "Open Sans", "Lato", "Montserrat", "Oswald", "Source Sans Pro",
    "Raleway", "PT Sans", "Merriweather", "Nunito", "Playfair Display", "Rubik",
    "Ubuntu", "Poppins", "Mukta", "Work Sans", "Noto Sans", "Fira Sans",
    "Quicksand", "Titillium Web", "Heebo", "Barlow", "Cabin", "Exo 2",
    "Josefin Sans", "Dosis", "Karla", "Inconsolata", "Hind", "Nanum Gothic",
    "Anton", "Bebas Neue", "Cinzel", "Cormorant Garamond", "Dancing Script",
    "Comfortaa", "Pacifico", "Sacramento", "Lobster", "Righteous", "Satisfy",
    "Great Vibes", "Courgette", "Caveat", "Shadows Into Light", "Permanent Marker",
    "Kaushan Script", "Amatic SC", "Architects Daughter", "Bad Script", "Cookie",
    "Handlee", "Marck Script", "Pinyon Script", "Rochester", "Ruthie",
    "EB Garamond", "Libre Baskerville", "Crimson Text", "Cardo", "Spectral",
    "Gentium Book Basic", "Lora", "Bitter", "Arvo", "Zilla Slab", "Rokkitt",
    "Domine", "Volkhov", "Neuton", "Glegoo", "Enriqueta", "Libre Caslon Text",
    "Abril Fatface", "Alfa Slab One", "Black Ops One", "Fjalla One", "Graduate",
    "Lilita One", "Luckiest Guy", "Passion One", "Permanent Marker", "Russo One",
    "Baloo 2", "Bree Serif", "Crete Round", "Della Respira", "Lustria", "Neucha",
    "Noticia Text", "Old Standard TT", "Poly", "Tinos", "Trirong", "Vesper Libre",
    "Exo", "Michroma", "Nova Mono", "Orbitron", "Rajdhani", "Share Tech Mono",
    "Space Mono", "VT323", "Audiowide", "Electrolize", "Nova Square", "Quantico",
    "Source Code Pro", "Anonymous Pro", "Cousine", "Cutive Mono", "PT Mono",
    "IBM Plex Mono", "Fira Code", "JetBrains Mono", "Overpass Mono", "Oxygen Mono",
    "Noto Serif", "Lato", "Alegreya", "Cormorant", "Gilda Display", "Italiana",
    "Josefin Slab", "Libre Bodoni", "Marcellus", "Mrs Saint Delafield", "Oleo Script",
    "Philosopher", "Poiret One", "Proza Libre", "Tenor Sans", "Unna", "Vidaloka",
    "Yantramanav", "Yeseva One", "Zeyada", "Merienda", "Amita", "Yellowtail",
    "Allura", "Alex Brush", "Parisienne", "Kristi", "Euphoria Script", "Forum",
    "Adamina", "Alegreya Sans", "Asap", "Barlow Condensed", "Barlow Semi Condensed",
    "Cairo", "Catamaran", "Chivo", "Didact Gothic", "Encode Sans", "Fira Sans Condensed",
    "Hind Madurai", "Hind Siliguri", "Kanit", "Mada", "Manrope", "Mulish", "Noto Sans JP",
    "Overpass", "Play", "Prompt", "Public Sans", "Saira", "Signika", "Sora", "Spartan",
    "Varela Round", "Yanone Kaffeesatz", "Zilla Slab Highlight", "Inter",
].sort();

// ============================================================
// State
// ============================================================
const state = {
    templateUrl: null,
    templateNaturalW: 0,
    templateNaturalH: 0,
    canvasDisplayW: 0,
    canvasDisplayH: 0,
    columns: [],
    fields: [],
    customFonts: [],          // [{name, filename, path, url}]
    loadedGFonts: new Set(),  // Google Font names already injected into <head>
    generatedCount: 0,
    totalRows: 0,
    currentStep: 1,
};

// ============================================================
// Fabric.js canvas instance
// ============================================================
let canvas = null;

// ============================================================
// Utility helpers
// ============================================================

function toast(message, type = "info", duration = 4000) {
    const el = document.getElementById("toast");
    el.textContent = message;
    el.className = `show ${type}`;
    clearTimeout(el._timer);
    el._timer = setTimeout(() => { el.className = ""; }, duration);
}

function setLoading(btnId, loading) {
    const btn = document.getElementById(btnId);
    if (!btn) return;
    btn.disabled = loading;
    btn.dataset.originalText = btn.dataset.originalText || btn.innerHTML;
    btn.innerHTML = loading
        ? `<span>⏳</span> Working…`
        : btn.dataset.originalText;
}

// ============================================================
// Step navigation
// ============================================================

function goToStep(step) {
    if (step < 1 || step > 5) return;
    if (step > state.currentStep + 1) {
        toast("Please complete the current step first.", "error");
        return;
    }
    state.currentStep = step;

    document.querySelectorAll(".step-pill").forEach(pill => {
        const n = parseInt(pill.dataset.step);
        pill.classList.remove("active", "done");
        if (n === step) pill.classList.add("active");
        else if (n < step) pill.classList.add("done");
    });

    document.querySelectorAll(".step-section").forEach(sec => {
        sec.classList.toggle("active", parseInt(sec.dataset.step) === step);
    });

    if (step === 3) initEditorCanvas();
    if (step === 4) {
        const el = document.getElementById("info-fields-count");
        if (el) el.textContent = state.fields.length;
    }
}

// ============================================================
// Step 1 — Template Upload
// ============================================================

function initTemplateUpload() {
    const zone = document.getElementById("template-zone");
    const input = document.getElementById("template-input");

    zone.addEventListener("dragover", e => { e.preventDefault(); zone.classList.add("dragover"); });
    zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
    zone.addEventListener("drop", e => {
        e.preventDefault(); zone.classList.remove("dragover");
        if (e.dataTransfer.files.length) handleTemplateFile(e.dataTransfer.files[0]);
    });
    input.addEventListener("change", () => {
        if (input.files.length) handleTemplateFile(input.files[0]);
    });
}

async function handleTemplateFile(file) {
    if (!file.type.match(/image\/(jpeg|jpg|png)/i)) {
        toast("Please upload a PNG or JPEG image.", "error"); return;
    }
    const fd = new FormData();
    fd.append("template", file);
    setLoading("btn-next-1", true);
    try {
        const res = await fetch("/upload-template", { method: "POST", body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);

        state.templateUrl = data.url;
        const img = document.getElementById("template-preview-img");
        img.src = data.url;
        img.onload = () => {
            state.templateNaturalW = img.naturalWidth;
            state.templateNaturalH = img.naturalHeight;
        };
        document.getElementById("template-preview-wrap").style.display = "block";
        document.getElementById("btn-next-1").disabled = false;
        toast("Template uploaded!", "success");
    } catch (err) {
        toast(`Upload failed: ${err.message}`, "error");
    } finally {
        setLoading("btn-next-1", false);
    }
}

// ============================================================
// Step 2 — Excel Upload
// ============================================================

function initExcelUpload() {
    const zone = document.getElementById("excel-zone");
    const input = document.getElementById("excel-input");

    zone.addEventListener("dragover", e => { e.preventDefault(); zone.classList.add("dragover"); });
    zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
    zone.addEventListener("drop", e => {
        e.preventDefault(); zone.classList.remove("dragover");
        if (e.dataTransfer.files.length) handleExcelFile(e.dataTransfer.files[0]);
    });
    input.addEventListener("change", () => {
        if (input.files.length) handleExcelFile(input.files[0]);
    });
}

async function handleExcelFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        toast("Please upload an .xlsx Excel file.", "error"); return;
    }
    const fd = new FormData();
    fd.append("excel", file);
    setLoading("btn-next-2", true);
    try {
        const res = await fetch("/upload-excel", { method: "POST", body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);

        state.columns = data.columns;
        const tags = document.getElementById("column-tags");
        tags.innerHTML = data.columns.map(c => `<span class="column-tag">${c}</span>`).join("");
        document.getElementById("excel-columns-wrap").style.display = "block";
        document.getElementById("btn-next-2").disabled = false;
        toast(`Found ${data.columns.length} columns`, "success");
    } catch (err) {
        toast(`Upload failed: ${err.message}`, "error");
    } finally {
        setLoading("btn-next-2", false);
    }
}

// ============================================================
// Font loading helpers
// ============================================================

/**
 * Inject a Google Font into the page <head> so Fabric.js can render it.
 * Uses the Google Fonts CSS2 API which supports wght axis for bold.
 */
function loadGoogleFont(fontName) {
    if (state.loadedGFonts.has(fontName)) return Promise.resolve();
    return new Promise(resolve => {
        const link = document.createElement("link");
        link.rel = "stylesheet";
        link.href = `https://fonts.googleapis.com/css2?family=${encodeURIComponent(fontName)}:ital,wght@0,400;0,700;1,400;1,700&display=swap`;
        link.onload = () => { state.loadedGFonts.add(fontName); resolve(); };
        link.onerror = () => resolve(); // fail silently
        document.head.appendChild(link);
    });
}

/**
 * Inject an uploaded custom font into the page via @font-face
 * so Fabric.js can display it on the canvas.
 */
function loadCustomFont(name, url) {
    if (state.loadedGFonts.has(name)) return Promise.resolve();
    return new Promise(resolve => {
        const style = document.createElement("style");
        style.textContent = `@font-face { font-family: "${name}"; src: url("${url}"); }`;
        document.head.appendChild(style);
        // Give browser a moment to register the font
        setTimeout(() => { state.loadedGFonts.add(name); resolve(); }, 300);
    });
}

// ============================================================
// Custom font file upload (Step 3 sidebar)
// ============================================================

async function handleCustomFontUpload(file) {
    if (!file.name.match(/\.(ttf|otf|woff|woff2)$/i)) {
        toast("Please upload a .ttf or .otf font file.", "error"); return;
    }
    const fd = new FormData();
    fd.append("font", file);
    try {
        const res = await fetch("/upload-font", { method: "POST", body: fd });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);

        state.customFonts.push(data);
        await loadCustomFont(data.name, data.url);
        renderFontSearchDropdowns();   // rebuild dropdowns in all fields
        updateCustomFontTagList();     // update sidebar tag list
        toast(`Font "${data.name}" uploaded & ready!`, "success");
    } catch (err) {
        toast(`Font upload failed: ${err.message}`, "error");
    }
}

/** Render the list of uploaded custom fonts as tags in the sidebar */
function updateCustomFontTagList() {
    const container = document.getElementById("custom-font-list");
    if (!container) return;
    if (!state.customFonts.length) {
        container.innerHTML = "";
        return;
    }
    container.innerHTML = state.customFonts
        .map(f => `<span class="custom-font-tag">🔤 ${f.name}</span>`)
        .join("");
}


/** Load existing uploaded fonts from the server on page load */
async function loadExistingCustomFonts() {
    try {
        const res = await fetch("/fonts");
        const data = await res.json();
        for (const font of data.fonts || []) {
            state.customFonts.push(font);
            await loadCustomFont(font.name, font.url);
        }
    } catch (_) { /* server not ready — ignore */ }
}

// ============================================================
// Step 3 — Fabric.js Editor
// ============================================================

const CANVAS_MAX_W = 820;

function initEditorCanvas() {
    if (!state.templateUrl) return;

    const aspectRatio = state.templateNaturalH / state.templateNaturalW;
    const displayW = Math.min(CANVAS_MAX_W, state.templateNaturalW);
    const displayH = Math.round(displayW * aspectRatio);
    state.canvasDisplayW = displayW;
    state.canvasDisplayH = displayH;

    if (canvas) { canvas.dispose(); canvas = null; }

    const wrap = document.getElementById("canvas-wrap");
    wrap.innerHTML = `<canvas id="editor-canvas"></canvas>`;

    canvas = new fabric.Canvas("editor-canvas", {
        width: displayW, height: displayH,
        selection: true, backgroundColor: "#222",
    });

    fabric.Image.fromURL(state.templateUrl, img => {
        img.scaleToWidth(displayW);
        img.scaleToHeight(displayH);
        img.set({ selectable: false, evented: false });
        canvas.setBackgroundImage(img, canvas.renderAll.bind(canvas));
    }, { crossOrigin: "anonymous" });

    canvas.on("selection:created", updateFieldHighlight);
    canvas.on("selection:updated", updateFieldHighlight);
    canvas.on("selection:cleared", () => {
        document.querySelectorAll(".field-item").forEach(el => el.classList.remove("selected"));
    });

    renderFieldList();
}

function updateFieldHighlight(e) {
    const obj = e.selected?.[0];
    if (!obj) return;
    document.querySelectorAll(".field-item").forEach(item => {
        item.classList.toggle("selected", item.dataset.id === obj._fieldId);
    });
}

// ============================================================
// Field management
// ============================================================

let _fieldIdCounter = 0;

function addField() {
    if (!canvas) { toast("Go to Step 3 to add fields.", "error"); return; }
    if (!state.columns.length) { toast("Please upload an Excel file first.", "error"); return; }

    const fieldId = String(++_fieldIdCounter);
    const defaultCol = state.columns[0];
    const defaultFont = "Montserrat";
    const defaultSize = 40;
    const defaultColor = "#ffffff";

    // Pre-load the default Google Font
    loadGoogleFont(defaultFont).then(() => {
        const text = new fabric.IText(defaultCol, {
            left: state.canvasDisplayW / 2,
            top: state.canvasDisplayH / 2,
            originX: "center",
            originY: "center",
            fontSize: defaultSize,
            fill: defaultColor,
            fontFamily: defaultFont,
            textAlign: "center",
            fontWeight: "normal",
            fontStyle: "normal",
            editable: false,
        });
        text._fieldId = fieldId;

        canvas.add(text);
        canvas.setActiveObject(text);
        canvas.renderAll();

        state.fields.push({
            id: fieldId,
            column: defaultCol,
            fontFamily: defaultFont,
            fontPath: null,      // resolved on backend; null = Google Font (downloaded separately)
            fontSize: defaultSize,
            color: defaultColor,
            align: "center",
            bold: false,
            italic: false,
            _fabricObj: text,
        });

        renderFieldList();
        toast(`Added field: "${defaultCol}"`, "info");
    });
}

function removeField(fieldId) {
    const idx = state.fields.findIndex(f => f.id === fieldId);
    if (idx === -1) return;
    canvas.remove(state.fields[idx]._fabricObj);
    state.fields.splice(idx, 1);
    renderFieldList();
    canvas.renderAll();
}

async function updateFieldProperty(fieldId, property, value) {
    const field = state.fields.find(f => f.id === fieldId);
    if (!field) return;

    // Convert checkbox values
    if (property === "bold" || property === "italic") value = Boolean(value);
    field[property] = value;

    const obj = field._fabricObj;

    if (property === "column") {
        obj.set("text", value);
    } else if (property === "fontFamily") {
        // Check if it's a custom uploaded font
        const custom = state.customFonts.find(f => f.name === value);
        if (custom) {
            field.fontPath = custom.path;   // absolute server path; backend will use this
            await loadCustomFont(custom.name, custom.url);
        } else {
            field.fontPath = null;           // Google Font — backend downloads via fontPath=null
            await loadGoogleFont(value);
        }
        obj.set("fontFamily", value);
    } else if (property === "fontSize") {
        obj.set("fontSize", parseInt(value) || 30);
    } else if (property === "color") {
        obj.set("fill", value);
    } else if (property === "align") {
        obj.set("textAlign", value);
    } else if (property === "bold") {
        obj.set("fontWeight", value ? "bold" : "normal");
    } else if (property === "italic") {
        obj.set("fontStyle", value ? "italic" : "normal");
    }

    canvas.renderAll();
}

// ============================================================
// Font search dropdown builder
// ============================================================

/** Build a searchable <datalist> + <input> combo for font selection */
function buildFontSelector(field) {
    const allFonts = [
        ...state.customFonts.map(f => ({ label: `⭐ ${f.name} (custom)`, value: f.name })),
        ...GOOGLE_FONTS.map(f => ({ label: f, value: f })),
    ];

    const datalistId = `font-list-${field.id}`;
    const options = allFonts.map(f =>
        `<option value="${f.value}">${f.label}</option>`
    ).join("");

    return `
    <label>Font Family</label>
    <div class="font-search-wrap">
      <input
        type="text"
        list="${datalistId}"
        id="font-input-${field.id}"
        value="${field.fontFamily}"
        placeholder="Search fonts…"
        autocomplete="off"
        oninput="handleFontInput('${field.id}', this.value)"
        onchange="handleFontChange('${field.id}', this.value)"
      />
      <datalist id="${datalistId}">${options}</datalist>
    </div>
  `;
}

function handleFontInput(fieldId, value) {
    // Live preview: if value is in GOOGLE_FONTS or custom, load + apply
    const match = GOOGLE_FONTS.find(f => f.toLowerCase() === value.toLowerCase())
        || state.customFonts.find(f => f.name.toLowerCase() === value.toLowerCase())?.name;
    if (match) handleFontChange(fieldId, match);
}

function handleFontChange(fieldId, value) {
    const exact = GOOGLE_FONTS.find(f => f === value)
        || state.customFonts.find(f => f.name === value)?.name;
    if (exact) updateFieldProperty(fieldId, "fontFamily", exact);
}

/** Re-render only the font selector inside existing field items (avoids full re-render) */
function renderFontSearchDropdowns() {
    state.fields.forEach(f => {
        const wrap = document.getElementById(`font-wrap-${f.id}`);
        if (wrap) wrap.innerHTML = buildFontSelector(f);
    });
}

// ============================================================
// Render the sidebar field list
// ============================================================

function renderFieldList() {
    const list = document.getElementById("field-list");
    if (!list) return;

    if (!state.fields.length) {
        list.innerHTML = `<p style="color:var(--text-muted);font-size:0.8rem;text-align:center">
      No fields yet. Click "Add Field" to start.
    </p>`;
        return;
    }

    list.innerHTML = state.fields.map(f => `
    <div class="field-item" data-id="${f.id}">
      <div class="field-header">
        <span class="field-name">${f.column}</span>
        <button class="btn btn-danger" style="padding:3px 10px;font-size:0.75rem"
          onclick="removeField('${f.id}')">✕</button>
      </div>

      <label>Excel Column</label>
      <select onchange="updateFieldProperty('${f.id}', 'column', this.value)">
        ${state.columns.map(c =>
        `<option value="${c}" ${c === f.column ? "selected" : ""}>${c}</option>`
    ).join("")}
      </select>

      <!-- Font search -->
      <div id="font-wrap-${f.id}">${buildFontSelector(f)}</div>

      <div class="field-row">
        <div>
          <label>Font Size</label>
          <input type="number" min="8" max="300" value="${f.fontSize}"
            onchange="updateFieldProperty('${f.id}', 'fontSize', this.value)">
        </div>
        <div>
          <label>Color</label>
          <input type="color" value="${f.color}"
            oninput="updateFieldProperty('${f.id}', 'color', this.value)">
        </div>
      </div>

      <label>Alignment</label>
      <select onchange="updateFieldProperty('${f.id}', 'align', this.value)">
        <option value="left"   ${f.align === "left" ? "selected" : ""}>Left</option>
        <option value="center" ${f.align === "center" ? "selected" : ""}>Center</option>
        <option value="right"  ${f.align === "right" ? "selected" : ""}>Right</option>
      </select>

      <!-- Bold / Italic -->
      <div class="style-toggles">
        <button class="style-btn ${f.bold ? "active" : ""}" title="Bold"
          onclick="toggleStyle('${f.id}','bold',  this)"><b>B</b></button>
        <button class="style-btn ${f.italic ? "active" : ""}" title="Italic"
          onclick="toggleStyle('${f.id}','italic',this)"><i>I</i></button>
      </div>
    </div>
  `).join("");
}

function toggleStyle(fieldId, property, btn) {
    const field = state.fields.find(f => f.id === fieldId);
    if (!field) return;
    const newVal = !field[property];
    updateFieldProperty(fieldId, property, newVal);
    btn.classList.toggle("active", newVal);
}

// ============================================================
// Build payload to send to backend
// ============================================================

function buildFieldsPayload() {
    const scaleX = state.templateNaturalW / state.canvasDisplayW;
    const scaleY = state.templateNaturalH / state.canvasDisplayH;

    return state.fields.map(f => {
        const obj = f._fabricObj;
        return {
            column: f.column,
            x: obj.left * scaleX,
            y: obj.top * scaleY,
            fontSize: Math.round(f.fontSize * Math.max(scaleX, scaleY)),
            color: f.color,
            align: f.align,
            fontFamily: f.fontFamily,
            fontPath: f.fontPath || null,   // null = Google font (backend will use system fallback)
            bold: f.bold,
            italic: f.italic,
        };
    });
}

// ============================================================
// Step 4 — Preview & Generate
// ============================================================

async function previewCertificate() {
    if (!validateReadyToGenerate()) return;
    setLoading("btn-preview", true);
    try {
        const res = await fetch("/preview", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fields: buildFieldsPayload() }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);

        document.getElementById("preview-modal-img").src = data.url;
        document.getElementById("preview-modal").classList.add("open");
    } catch (err) {
        toast(`Preview failed: ${err.message}`, "error");
    } finally {
        setLoading("btn-preview", false);
    }
}

function closePreview() {
    document.getElementById("preview-modal").classList.remove("open");
}

async function generateAll() {
    if (!validateReadyToGenerate()) return;
    setLoading("btn-generate", true);
    showProgress(true);
    try {
        const res = await fetch("/generate", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fields: buildFieldsPayload() }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);

        state.generatedCount = data.generated;
        state.totalRows = data.total;
        setProgressFill(100);
        toast(`Generated ${data.generated} / ${data.total} certificates!`, "success");
        if (data.errors?.length) {
            console.warn("Row errors:", data.errors);
            toast(`${data.errors.length} rows had errors (see console).`, "error");
        }
        updateDownloadStep();
        setTimeout(() => goToStep(5), 800);
    } catch (err) {
        toast(`Generation failed: ${err.message}`, "error");
        showProgress(false);
    } finally {
        setLoading("btn-generate", false);
    }
}

function validateReadyToGenerate() {
    if (!state.templateUrl) { toast("No template uploaded.", "error"); return false; }
    if (!state.columns.length) { toast("No Excel file uploaded.", "error"); return false; }
    if (!state.fields.length) { toast("No fields placed on the canvas.", "error"); return false; }
    return true;
}

function showProgress(show) {
    const wrap = document.getElementById("progress-wrap");
    if (wrap) wrap.style.display = show ? "block" : "none";
    if (show) setProgressFill(30);
}

function setProgressFill(pct) {
    const fill = document.getElementById("progress-fill");
    if (fill) fill.style.width = `${pct}%`;
    const label = document.getElementById("progress-label");
    if (label) label.textContent = pct < 100 ? "Generating certificates…" : "Done!";
}

// ============================================================
// Step 5 — Download
// ============================================================

function updateDownloadStep() {
    document.getElementById("stat-generated").textContent = state.generatedCount;
    document.getElementById("stat-total").textContent = state.totalRows;
}

function downloadZip() { window.location.href = "/download-zip"; }

// ============================================================
// Save Configuration JSON
// ============================================================

function saveConfig() {
    if (!state.fields.length) { toast("No fields to save.", "error"); return; }
    const config = { templateUrl: state.templateUrl, columns: state.columns, fields: buildFieldsPayload() };
    const blob = new Blob([JSON.stringify(config, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "certificate_config.json";
    a.click();
    toast("Configuration saved!", "success");
}

// ============================================================
// Initialise on DOM ready
// ============================================================

document.addEventListener("DOMContentLoaded", async () => {
    initTemplateUpload();
    initExcelUpload();

    // Custom font upload input
    const fontInput = document.getElementById("custom-font-input");
    if (fontInput) {
        fontInput.addEventListener("change", () => {
            if (fontInput.files.length) handleCustomFontUpload(fontInput.files[0]);
        });
    }
    // Font upload zone drag-and-drop
    const fontZone = document.getElementById("font-upload-zone");
    if (fontZone) {
        fontZone.addEventListener("dragover", e => { e.preventDefault(); fontZone.classList.add("dragover"); });
        fontZone.addEventListener("dragleave", () => fontZone.classList.remove("dragover"));
        fontZone.addEventListener("drop", e => {
            e.preventDefault(); fontZone.classList.remove("dragover");
            if (e.dataTransfer.files.length) handleCustomFontUpload(e.dataTransfer.files[0]);
        });
    }

    // Step pill navigation
    document.querySelectorAll(".step-pill").forEach(pill => {
        pill.addEventListener("click", () => goToStep(parseInt(pill.dataset.step)));
    });

    // Modal close
    document.getElementById("preview-modal").addEventListener("click", e => {
        if (e.target === e.currentTarget) closePreview();
    });
    document.addEventListener("keydown", e => { if (e.key === "Escape") closePreview(); });

    // Load any previously uploaded custom fonts
    await loadExistingCustomFonts();

    goToStep(1);
});
