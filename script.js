'use strict';

/* ==========================================================
   STATE
   ========================================================== */

// Source frame count is now computed from range duration × target fps.
// These are guardrails to keep memory sane.
const MIN_SOURCE_FRAMES = 30;     // below this, scrubbing feels choppy
const MAX_SOURCE_FRAMES = 1800;   // 60s × 30fps — hard ceiling

const MAX_FILE_BYTES = 200 * 1024 * 1024;
const MAX_CANVAS_DIM = 1000;

const state = {
    video: { width: 0, height: 0, duration: 0, name: '', size: 0, blobUrl: null, probe: null },
    range: { in: 0, out: 0 },
    totalFrames: 0,
    frames: [],
    tiles: [],
    hoverTile: null,
    activeTile: null,
    isDragging: false,
    dragStartX: 0,
    dragStartFrame: 0,
    isExporting: false,
    scrubbingTile: null,
    ready: false
};

const anim = {
    mode: 'standard',
    raf: null,
    lastNow: 0,
    elapsed: 0,
    tileOffsets: [],
    tilePhases: [],
    _loop: null,
    playing: false,
    // Loop mode: 'wrap' | 'pingpong' | 'hold'
    loopMode: 'wrap',
    // Ping-pong direction tracking
    pingpongForward: true,
    // Stutter factor (1-6)
    stutter: 1,
};

/* ==========================================================
   DOM
   ========================================================== */

const $ = id => document.getElementById(id);
const canvas = $('mosaicCanvas');
const ctx = canvas.getContext('2d');

const el = {
    sourceCard: $('sourceCard'),
    videoUpload: $('videoUpload'),
    sectionRange: $('sectionRange'),
    rangeSelector: $('rangeSelector'),
    rangeStrip: $('rangeStrip'),
    rangeMaskL: $('rangeMaskL'),
    rangeMaskR: $('rangeMaskR'),
    rangeActive: $('rangeActive'),
    rangeHandleL: $('rangeHandleL'),
    rangeHandleR: $('rangeHandleR'),
    rangeIn: $('rangeIn'),
    rangeOut: $('rangeOut'),
    rangeDur: $('rangeDur'),
    inputCols: $('inputCols'),
    inputRows: $('inputRows'),
    inputDuration: $('inputDuration'),
    durationValue: $('durationValue'),
    checkSquare: $('checkSquare'),
    checkGrid: $('checkGrid'),
    checkLoop: $('checkLoop'),
    selectRes: $('selectRes'),
    selectFps: $('selectFps'),
    btnExportPng: $('btnExportPng'),
    btnExportVideo: $('btnExportVideo'),
    selectMode: $('selectMode'),
    inputSpatialAmt: $('inputSpatialAmt'),
    spatialAmtValue: $('spatialAmtValue'),
    spatialAmtRow: $('spatialAmtRow'),
    btnPlayPause: $('btnPlayPause'),
    playpauseIcon: $('playpauseIcon'),
    playpauseLabel: $('playpauseLabel'),
    checkSpatialShuffle: $('checkSpatialShuffle'),
    selectPattern: $('selectPattern'),
    patternAmtRow: $('patternAmtRow'),
    inputPatternAmt: $('inputPatternAmt'),
    patternAmtValue: $('patternAmtValue'),
    btnPatternApply: $('btnPatternApply'),
    btnPatternClear: $('btnPatternClear'),
    btnBrowse: $('btnBrowse'),
    emptyState: $('emptyState'),
    loadingOverlay: $('loadingOverlay'),
    loadingText: (() => { const el = $('loadingText'); return el && el.parentElement ? el : null; })(),
    progressFill: $('progressFill'),
    timeline: $('timeline'),
    tlTile: $('tlTile'),
    tlTime: $('tlTime'),
    tlFill: $('tlFill'),
    tlHead: $('tlHead'),
    statusCanvas: $('statusCanvas'),
    statusTiles: $('statusTiles'),
    statusPinned: $('statusPinned'),
    statusState: $('statusState'),
    toast: $('toast'),
    toastMsg: $('toastMsg'),
    confirmDialog: $('confirmDialog'),
    confirmTitle: $('confirmTitle'),
    confirmMsg: $('confirmMsg'),
    confirmOk: $('confirmOk'),
    confirmCancel: $('confirmCancel'),
    selectLoop: $('selectLoop'),
    inputStutter: $('inputStutter'),
    stutterValue: $('stutterValue'),
};

/* ==========================================================
   UTILITIES
   ========================================================== */

function showToast(msg, type = 'error') {
    el.toastMsg.innerText = msg;
    el.toast.classList.remove('info');
    if (type === 'info') el.toast.classList.add('info');
    el.toast.classList.add('show');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => el.toast.classList.remove('show'), 3600);
}

function formatTime(sec) {
    if (!isFinite(sec)) return '00:00.00';
    const m = Math.floor(sec / 60);
    const s = sec - m * 60;
    return `${String(m).padStart(2, '0')}:${s.toFixed(2).padStart(5, '0')}`;
}

function formatBytes(b) {
    if (b < 1024 * 1024) return (b / 1024).toFixed(1) + ' KB';
    return (b / (1024 * 1024)).toFixed(1) + ' MB';
}

function confirm(title, msg, okLabel = 'Confirm') {
    return new Promise(resolve => {
        el.confirmTitle.innerText = title;
        el.confirmMsg.innerText = msg;
        el.confirmOk.innerText = okLabel;
        el.confirmDialog.classList.add('visible');
        const cleanup = (v) => {
            el.confirmDialog.classList.remove('visible');
            el.confirmOk.onclick = null;
            el.confirmCancel.onclick = null;
            resolve(v);
        };
        el.confirmOk.onclick = () => cleanup(true);
        el.confirmCancel.onclick = () => cleanup(false);
    });
}

/* ==========================================================
   FILE HANDLING
   ========================================================== */

/* ==========================================================
   DRAG & DROP
   ========================================================== */

// Browser default for a dropped file is to navigate to it.
// To opt out, we MUST call preventDefault() on BOTH 'dragover' AND 'drop'
// — at the element level, not just on window. The HTML5 drag spec says
// "if dragover was not prevented, the drop is not allowed" and the browser
// falls back to its default (opening the file in a new tab).

// Global safety net: prevent browser from handling any drag that escapes our dropzones.
// Must prevent on dragover too, not just drop.
['dragenter', 'dragover', 'drop'].forEach(n => {
    window.addEventListener(n, (e) => {
        // Only prevent if it's a file drag — don't interfere with internal UI drags
        if (e.dataTransfer && e.dataTransfer.types && e.dataTransfer.types.includes('Files')) {
            e.preventDefault();
        }
    });
});

function bindDropZone(node) {
    node.addEventListener('dragenter', (e) => {
        e.preventDefault();
        node.classList.add('dragover');
    });
    node.addEventListener('dragover', (e) => {
        // CRITICAL: preventDefault here signals "I accept this drop".
        // Without it, browser opens the file in a new tab.
        e.preventDefault();
        if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
    });
    node.addEventListener('dragleave', (e) => {
        // Only remove the class if we actually left the element (not a child)
        if (e.currentTarget === e.target) node.classList.remove('dragover');
    });
    node.addEventListener('drop', (e) => {
        e.preventDefault();
        node.classList.remove('dragover');
        const f = e.dataTransfer?.files?.[0];
        if (f) handleFile(f);
    });
}

bindDropZone($('canvasWrap'));
bindDropZone(el.sourceCard);

el.sourceCard.addEventListener('click', () => el.videoUpload.click());
el.btnBrowse.addEventListener('click', () => el.videoUpload.click());
el.videoUpload.addEventListener('change', (e) => {
    const f = e.target.files?.[0];
    if (f) handleFile(f);
    e.target.value = ''; // allow re-selecting the same file
});

function handleFile(file) {
    if (!file) return;
    if (file.size > MAX_FILE_BYTES) {
        showToast(`File too large (${formatBytes(file.size)}, max 200 MB)`);
        return;
    }
    // Some OSes report an empty MIME type — fall back to the extension.
    const typeOk = /^video\/(mp4|webm|quicktime|x-m4v)$/.test(file.type);
    const extOk = !file.type && /\.(mp4|m4v|webm|mov)$/i.test(file.name);
    if (!typeOk && !extOk) {
        showToast(`Unsupported format: ${file.type || 'unknown'}`);
        return;
    }
    if (state.isExporting) {
        showToast('Wait for the export to finish before loading a new video.');
        return;
    }

    // Cancel any in-flight extraction and free the previous video's resources
    cancelExtraction();
    pause();
    releaseFrames();
    releaseProbe();
    if (state.video.blobUrl) URL.revokeObjectURL(state.video.blobUrl);

    state.ready = false;
    state.tiles = [];
    state.hoverTile = null;
    state.activeTile = null;
    state.video.name = file.name;
    state.video.size = file.size;
    state.video.blobUrl = URL.createObjectURL(file);

    loadVideo(state.video.blobUrl);
}

function releaseFrames() {
    for (const f of state.frames) {
        if (f && typeof f.close === 'function') f.close();
    }
    state.frames = [];
}

function releaseProbe() {
    const probe = state.video.probe;
    if (!probe) return;
    probe.onloadedmetadata = null;
    probe.onerror = null;
    probe.removeAttribute('src');
    probe.load();
    state.video.probe = null;
}

function loadVideo(src) {
    const probe = document.createElement('video');
    probe.preload = 'metadata';
    probe.muted = true;
    probe.src = src;
    state.video.probe = probe;
    let settled = false;

    const timeout = setTimeout(() => {
        if (settled) return;
        settled = true;
        if (state.video.probe === probe) releaseProbe();
        showToast('Video metadata timeout. Try a different file.');
    }, 8000);

    probe.onloadedmetadata = () => {
        if (settled || state.video.probe !== probe) return;
        settled = true;
        clearTimeout(timeout);

        if (!isFinite(probe.duration) || probe.duration <= 0) {
            showToast('Video has invalid duration metadata.');
            return;
        }
        if (!probe.videoWidth || !probe.videoHeight) {
            showToast('No video track found in this file.');
            return;
        }

        state.video.width = probe.videoWidth;
        state.video.height = probe.videoHeight;
        state.video.duration = probe.duration;
        state.video.probe = probe;

        // Sensible default range: full video if ≤10s, else first 10s
        state.range.in = 0;
        state.range.out = Math.min(probe.duration, 10);

        // Sensible default export duration: match range duration (up to 10s)
        const rangeDur = state.range.out - state.range.in;
        const defaultExportDur = Math.min(10, Math.max(2, rangeDur));
        el.inputDuration.value = defaultExportDur;
        el.durationValue.innerText = defaultExportDur.toFixed(1);
        el.inputDuration.max = Math.min(60, probe.duration).toFixed(1);

        updateSourceCard();
        el.sectionRange.style.display = 'flex';
        renderRangeUI();
        buildThumbstrip(probe);  // fire-and-forget, cosmetic
        extractFrames(probe);
    };

    probe.onerror = () => {
        if (settled || state.video.probe !== probe) return;
        settled = true;
        clearTimeout(timeout);
        showToast('Could not read video file.');
    };
}

function updateSourceCard() {
    const v = state.video;
    el.sourceCard.classList.remove('empty');

    // Built with textContent: the file name is user-controlled and must never be parsed as HTML.
    const thumb = document.createElement('div');
    thumb.className = 'source-thumb';
    thumb.textContent = '▶';

    const meta = document.createElement('div');
    meta.className = 'source-meta';
    const name = document.createElement('div');
    name.className = 'source-name';
    name.textContent = v.name;
    name.title = v.name;
    const specs = document.createElement('div');
    specs.className = 'source-specs';
    specs.textContent = `${v.width}×${v.height} · ${formatBytes(v.size)} · ${v.duration.toFixed(1)}s`;
    meta.append(name, specs);

    const btn = document.createElement('button');
    btn.className = 'btn-change';
    btn.id = 'btnChangeVideo';
    btn.type = 'button';
    btn.textContent = 'Change';
    btn.onclick = (e) => {
        e.stopPropagation();
        el.videoUpload.click();
    };

    el.sourceCard.replaceChildren(thumb, meta, btn);
}

/* ==========================================================
   FRAME EXTRACTION — with timeout + single-shot seeked handler
   ========================================================== */

/* ==========================================================
   FRAME EXTRACTION — dynamic count based on range × fps
   ========================================================== */

function computeSourceFrameCount() {
    // Source frame count = range_duration × fps.
    // This is independent of export duration — duration only controls
    // how long the output video plays, not how many unique source frames
    // we have available.
    const rangeDur = Math.max(0.1, state.range.out - state.range.in);
    const fps = parseInt(el.selectFps.value, 10);

    let count = Math.round(rangeDur * fps);
    count = Math.max(MIN_SOURCE_FRAMES, Math.min(MAX_SOURCE_FRAMES, count));
    return count;
}

// Each extraction gets a token; starting a new one (or loading a new video)
// invalidates the previous loop so two extractions never write into the same list.
let extractionToken = 0;
let extractionRunning = false;
let lastExtractedRange = null;   // range the current frames were sampled from

function cancelExtraction() {
    extractionToken++;
}

/**
 * Extract source frames for the current range × fps.
 * keepProject: true when re-extracting for an already-built grid (range/fps change) —
 * tiles, pins and scrub offsets are kept and rescaled instead of rebuilding the grid.
 */
async function extractFrames(probe, { keepProject = false, prevRange = null } = {}) {
    const token = ++extractionToken;
    extractionRunning = true;
    const prevN = state.totalFrames;
    const targetCount = computeSourceFrameCount();
    const frames = [];
    showLoading(`EXTRACTING ${targetCount} FRAMES`);

    const temp = document.createElement('canvas');
    temp.width = state.video.width;
    temp.height = state.video.height;
    const tCtx = temp.getContext('2d');

    const rangeDur = state.range.out - state.range.in;
    const discard = () => {
        for (const f of new Set(frames)) f.close();
    };

    for (let i = 0; i < targetCount; i++) {
        if (token !== extractionToken) { discard(); return; }
        try {
            const t = state.range.in + (i / Math.max(1, targetCount - 1)) * rangeDur;
            frames.push(await seekAndCapture(probe, tCtx, temp, t));
        } catch (err) {
            if (token !== extractionToken) { discard(); return; }
            console.warn(`Frame ${i} failed:`, err);
            if (frames.length > 0) {
                frames.push(frames[frames.length - 1]);
            } else {
                discard();
                extractionRunning = false;
                // A failed re-extraction keeps the previous frames (and their range) usable
                state.ready = state.frames.length > 0 && state.tiles.length > 0;
                if (state.ready && lastExtractedRange) {
                    state.range = { ...lastExtractedRange };
                    renderRangeUI();
                }
                hideLoading();
                showToast('Frame extraction failed. Try a different video.');
                return;
            }
        }
        el.progressFill.style.width = `${((i + 1) / targetCount) * 100}%`;
    }
    if (token !== extractionToken) { discard(); return; }

    // Swap in the new frames only once extraction completed
    releaseFrames();
    state.frames = frames;
    state.totalFrames = targetCount;
    lastExtractedRange = { ...state.range };
    extractionRunning = false;

    hideLoading();
    canvas.style.display = 'block';
    el.emptyState.style.display = 'none';
    el.timeline.classList.add('visible');
    state.ready = true;
    enableControls();
    if (keepProject && state.tiles.length > 0) {
        remapProjectFrames(prevRange || { ...state.range }, prevN, targetCount);
    } else {
        initProject();
    }
    updateStatus();
}

/**
 * Keep pins/offsets after a re-extraction.
 * Pinned frames stay on the same moment of the video (clamped to the new range);
 * offsets and phases keep the same duration in seconds.
 */
function remapProjectFrames(prevRange, prevN, newN) {
    const prevDur = Math.max(0.1, prevRange.out - prevRange.in);
    const newDur = Math.max(0.1, state.range.out - state.range.in);
    const prevRate = Math.max(1, prevN - 1) / prevDur;
    const newRate = Math.max(1, newN - 1) / newDur;
    const k = newRate / prevRate;
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

    for (const t of state.tiles) {
        const time = prevRange.in + t.frameIndex / prevRate;
        t.frameIndex = clamp(Math.round((time - state.range.in) * newRate), 0, newN - 1);
        t.frameOffset = clamp(Math.round(t.frameOffset * k), -(newN - 1), newN - 1);
    }
    anim.tilePhases = anim.tilePhases.map(p => p * k);
    renderAll();
    updateTimeline();
}

/** Re-extract frames after range or fps change. */
function reextractFrames() {
    if (!state.video.probe || state.isExporting) return;
    const hadProject = state.tiles.length > 0;
    pause();
    state.ready = false;
    extractFrames(state.video.probe, { keepProject: hadProject, prevRange: lastExtractedRange });
}

function seekAndCapture(probe, tCtx, temp, targetTime) {
    return new Promise((resolve, reject) => {
        let resolved = false;

        const onSeeked = async () => {
            if (resolved) return;
            resolved = true;
            probe.removeEventListener('seeked', onSeeked);
            clearTimeout(t);
            try {
                tCtx.drawImage(probe, 0, 0);
                // ImageBitmap: GPU-ready handle, ~10× faster drawImage than HTML Image,
                // and no base64 encode/decode roundtrip.
                resolve(await createImageBitmap(temp));
            } catch (err) {
                reject(err);
            }
        };

        const t = setTimeout(() => {
            if (resolved) return;
            resolved = true;
            probe.removeEventListener('seeked', onSeeked);
            reject(new Error(`seek timeout at ${targetTime.toFixed(2)}s`));
        }, 4000);

        probe.addEventListener('seeked', onSeeked);
        probe.currentTime = targetTime;
    });
}

/* ==========================================================
   RANGE UI — thumbstrip + in/out handles
   ========================================================== */

async function buildThumbstrip(probe) {
    // Generate ~10 small thumbnails across the full video for the range strip.
    // This is cosmetic — runs in parallel with frame extraction.
    // Uses a tiny separate video element so we don't interfere with the main probe.
    const src = state.video.blobUrl;
    const vid = document.createElement('video');
    el.rangeStrip.style.backgroundImage = '';
    try {
        const THUMB_COUNT = 10;
        const THUMB_W = 80;
        const THUMB_H = 40;
        vid.muted = true;
        vid.preload = 'metadata';
        vid.src = src;
        await new Promise(r => {
            if (vid.readyState >= 1) r();
            else vid.onloadedmetadata = r;
        });

        const stripCanvas = document.createElement('canvas');
        stripCanvas.width = THUMB_W * THUMB_COUNT;
        stripCanvas.height = THUMB_H;
        const sctx = stripCanvas.getContext('2d');

        for (let i = 0; i < THUMB_COUNT; i++) {
            if (state.video.blobUrl !== src) return;  // a new video was loaded
            const t = (i / (THUMB_COUNT - 1)) * state.video.duration * 0.999;
            await new Promise((res, rej) => {
                const to = setTimeout(res, 2000);  // forgive missing frames
                vid.onseeked = () => {
                    clearTimeout(to);
                    sctx.drawImage(vid, i * THUMB_W, 0, THUMB_W, THUMB_H);
                    res();
                };
                vid.currentTime = t;
            });
        }
        if (state.video.blobUrl !== src) return;
        el.rangeStrip.style.backgroundImage = `url(${stripCanvas.toDataURL('image/jpeg', 0.6)})`;
        el.rangeStrip.style.backgroundSize = '100% 100%';
    } catch (e) {
        // Thumbstrip failure is non-fatal — just leave it empty
        console.warn('Thumbstrip failed:', e);
    } finally {
        vid.removeAttribute('src');
        vid.load();
    }
}

function renderRangeUI() {
    const dur = state.video.duration;
    if (dur <= 0) return;
    const pctIn = (state.range.in / dur) * 100;
    const pctOut = (state.range.out / dur) * 100;

    el.rangeHandleL.style.left = pctIn + '%';
    el.rangeHandleR.style.left = pctOut + '%';
    el.rangeActive.style.left = pctIn + '%';
    el.rangeActive.style.right = (100 - pctOut) + '%';
    el.rangeMaskL.style.width = pctIn + '%';
    el.rangeMaskR.style.width = (100 - pctOut) + '%';

    el.rangeIn.innerText = formatTime(state.range.in);
    el.rangeOut.innerText = formatTime(state.range.out);
    const rd = state.range.out - state.range.in;
    el.rangeDur.innerText = rd.toFixed(2) + 's';
}

// Drag handles for range selector
let rangeDragging = null;  // 'L' or 'R' or null
const MIN_RANGE_SEC = 0.5;

function rangeFromClientX(clientX) {
    const r = el.rangeSelector.getBoundingClientRect();
    const pct = Math.max(0, Math.min(1, (clientX - r.left) / r.width));
    return pct * state.video.duration;
}

let rangeAtDragStart = null;

[el.rangeHandleL, el.rangeHandleR].forEach((h, idx) => {
    h.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        if (state.isExporting || extractionRunning) return;
        rangeDragging = idx === 0 ? 'L' : 'R';
        rangeAtDragStart = { ...state.range };
        try { h.setPointerCapture(e.pointerId); } catch (_) {}
    });
});

window.addEventListener('pointermove', (e) => {
    if (!rangeDragging) return;
    const t = rangeFromClientX(e.clientX);
    if (rangeDragging === 'L') {
        state.range.in = Math.max(0, Math.min(state.range.out - MIN_RANGE_SEC, t));
    } else {
        state.range.out = Math.min(state.video.duration, Math.max(state.range.in + MIN_RANGE_SEC, t));
    }
    renderRangeUI();
});

const endRangeDrag = () => {
    if (!rangeDragging) return;
    rangeDragging = null;
    // Export may be longer than the range (time-dilation effect) — no auto-clamp.
    // Re-extract source frames only if the range actually moved.
    const prev = rangeAtDragStart;
    rangeAtDragStart = null;
    if (prev && (prev.in !== state.range.in || prev.out !== state.range.out)) {
        reextractFrames();
    }
};
window.addEventListener('pointerup', endRangeDrag);
window.addEventListener('pointercancel', endRangeDrag);

/* ==========================================================
   PROJECT / GRID
   ========================================================== */

function initProject() {
    const cols = Math.max(1, Math.min(30, parseInt(el.inputCols.value) || 1));
    const isSquare = el.checkSquare.checked;
    const vW = state.video.width, vH = state.video.height;
    const vRatio = vW / vH;

    let rows = isSquare
        ? Math.max(1, Math.round(cols / vRatio))
        : Math.max(1, Math.min(30, parseInt(el.inputRows.value) || 1));

    el.inputRows.value = rows;
    el.inputRows.disabled = isSquare;

    // Canvas sizing
    let cW, cH;
    if (!isSquare) {
        const s = Math.min(1, MAX_CANVAS_DIM / vW, MAX_CANVAS_DIM / vH);
        cW = Math.round(vW * s);
        cH = Math.round(vH * s);
    } else {
        const ts = MAX_CANVAS_DIM / Math.max(cols, rows);
        cW = Math.round(cols * ts);
        cH = Math.round(rows * ts);
    }
    canvas.width = cW;
    canvas.height = cH;

    // Centered non-destructive crop
    const cRatio = cW / cH;
    let crW = vW, crH = vH;
    if (cRatio > vRatio) crH = vW / cRatio;
    else crW = vH * cRatio;

    const sW = crW / cols, sH = crH / rows;
    const oX = (vW - crW) / 2, oY = (vH - crH) / 2;

    state.tiles = [];
    for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
            state.tiles.push({
                x: c * (cW / cols),
                y: r * (cH / rows),
                w: cW / cols,
                h: cH / rows,
                srcX: oX + c * sW,
                srcY: oY + r * sH,
                origSrcX: oX + c * sW,
                origSrcY: oY + r * sH,
                srcW: sW,
                srcH: sH,
                frameIndex: 0,
                frameOffset: 0,
                isPinned: false,
                id: r * cols + c
            });
        }
    }
    rememberGrid();
    setMode(anim.mode);
    if (el.checkSpatialShuffle.checked && state.ready) {
        applySpatialShuffle(parseInt(el.inputSpatialAmt.value));
    }
    updateStatus();
}

/* ==========================================================
   PIPELINE: Single Source of Truth for frame index computation
   ========================================================== */

/**
 * Compute the final frame index for a tile given the raw output frame number.
 * Pipeline order (mandatory):
 *   1. Pin check — if pinned, return tile.frameIndex (frozen)
 *   2. Stutter — quantize outputFrame to stutter blocks
 *   3. Mode — determine base index per mode (Linear, Shuffle, etc.)
 *   4. Loop — apply wrap, ping-pong, or hold
 */
function computeFrameIndex(tile, outputFrame, cols, N, rate, elapsed, tileOffsets, tilePhases, i, dt) {
    // Step 1: Pin check
    if (tile.isPinned) return tile.frameIndex;

    // Step 2: Stutter — quantize outputFrame to blocks
    // At 1x: advances normally. At 3x: every 3 output frames, jumps 3 source frames.
    // This creates a dreamy, staccato "step-printing" effect (Wong Kar-wai style).
    const stutterFactor = anim.stutter;
    const stutterFrame = Math.floor(outputFrame / stutterFactor) * stutterFactor;

    // Step 3: Mode — determine base index using the stutter-quantized frame
    let baseIndex;
    const t = elapsed;

    switch (anim.mode) {
        case 'standard': {
            const rawPhase = stutterFrame + tile.frameOffset;
            baseIndex = Math.floor(rawPhase);
            break;
        }
        case 'linear-lr': {
            const phase = stutterFrame;
            const col = i % cols;
            baseIndex = Math.round((phase + (col / Math.max(1, cols - 1)) * N * 0.6));
            break;
        }
        case 'linear-rl': {
            const phase = stutterFrame;
            const col = i % cols;
            baseIndex = Math.round((phase + ((cols - 1 - col) / Math.max(1, cols - 1)) * N * 0.6));
            break;
        }
        case 'temporal-shuffle': {
            const phase = stutterFrame;
            baseIndex = Math.round(phase + tileOffsets[i] * N);
            break;
        }
        case 'drunk': {
            const col = i % cols;
            const row = Math.floor(i / cols);
            const v = valueNoise3(
                col * 0.7 + tileOffsets[i] * 5.1,
                row * 0.7 + tileOffsets[i] * 3.7,
                t * 0.6
            );
            baseIndex = Math.round(v * (N - 1));
            break;
        }
        case 'perlin-flow': {
            const col = i % cols;
            const row = Math.floor(i / cols);
            const v = valueNoise3(col * 0.35, row * 0.35, t * 0.12) * 2 - 1;
            // Integrate with the real frame step so speed is independent of display refresh rate
            tilePhases[i] = ((tilePhases[i] + v * rate * dt) % N + N) % N;
            baseIndex = Math.round(tilePhases[i]);
            break;
        }
        default:
            baseIndex = 0;
    }

    // Step 4: Loop — apply wrap, ping-pong, or hold
    return applyLoop(baseIndex, N);
}

/**
 * Apply loop logic to a raw frame index.
 * @param {number} idx - Raw frame index (may be negative or exceed N-1)
 * @param {number} N - Total number of frames
 * @returns {number} Clamped frame index [0, N-1]
 */
function applyLoop(idx, N) {
    const loopMode = anim.loopMode;

    switch (loopMode) {
        case 'wrap':
            // Wrap (Loop): arrive at end → restart from 0
            return ((idx % N) + N) % N;

        case 'pingpong': {
            // Ping-pong: arrive at end → reverse direction (smooth bounce)
            const period = 2 * (N - 1);
            if (period <= 0) return 0;
            const mod = ((idx % period) + period) % period;
            return mod < N ? mod : period - mod;
        }

        case 'hold':
            // Hold last frame: freeze on last frame when past end
            return Math.max(0, Math.min(N - 1, idx));

        default:
            return ((idx % N) + N) % N;
    }
}

/* ==========================================================
   RENDER
   ========================================================== */

function renderAll() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    for (const t of state.tiles) {
        const img = state.frames[t.frameIndex];
        if (img) {
            ctx.drawImage(img, t.srcX, t.srcY, t.srcW, t.srcH, t.x, t.y, t.w, t.h);
        }
    }

    // Pin overlays — SKIPPED during export (bug fix: pins used to appear in PNG)
    if (!state.isExporting) {
        for (const t of state.tiles) {
            if (t.isPinned) {
                ctx.strokeStyle = 'rgba(255, 184, 0, 0.9)';
                ctx.lineWidth = 2;
                ctx.strokeRect(t.x + 1.5, t.y + 1.5, t.w - 3, t.h - 3);
                ctx.fillStyle = 'rgba(255, 184, 0, 1)';
                ctx.fillRect(t.x + t.w - 12, t.y + 4, 8, 8);
            }
        }

        // Active tile highlight
        if (state.activeTile) {
            const t = state.activeTile;
            ctx.strokeStyle = 'rgba(0, 255, 163, 0.9)';
            ctx.lineWidth = 2;
            ctx.strokeRect(t.x + 1, t.y + 1, t.w - 2, t.h - 2);
        }

        if (el.checkGrid.checked) drawGrid();
    }
}

function drawGrid() {
    // Magenta with subtle glow for high visibility against any content
    ctx.save();
    ctx.shadowColor = 'rgba(26, 31, 40, .51)';
    ctx.shadowBlur = 2;
    ctx.strokeStyle = 'rgba(26, 31, 40, 1)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    const cols = parseInt(el.inputCols.value);
    const rows = parseInt(el.inputRows.value);
    for (let i = 1; i < cols; i++) {
        ctx.moveTo(i * (canvas.width / cols), 0);
        ctx.lineTo(i * (canvas.width / cols), canvas.height);
    }
    for (let i = 1; i < rows; i++) {
        ctx.moveTo(0, i * (canvas.height / rows));
        ctx.lineTo(canvas.width, i * (canvas.height / rows));
    }
    ctx.stroke();
    ctx.restore();
}

/* ==========================================================
   STATUS / TIMELINE
   ========================================================== */

function updateStatus() {
    if (!state.ready) {
        el.statusCanvas.innerText = '—';
        el.statusTiles.innerText = '—';
        el.statusPinned.style.display = 'none';
        el.statusState.style.display = 'none';
        return;
    }
    el.statusCanvas.innerText = `CANVAS ${canvas.width}×${canvas.height}`;
    const cols = parseInt(el.inputCols.value);
    const rows = parseInt(el.inputRows.value);
    el.statusTiles.innerText = `${cols}×${rows} TILES`;

    const pinned = state.tiles.filter(t => t.isPinned).length;
    if (pinned > 0) {
        el.statusPinned.style.display = 'inline';
        el.statusPinned.innerText = `${pinned} PINNED`;
    } else {
        el.statusPinned.style.display = 'none';
    }
    el.statusState.style.display = 'inline-flex';
    el.statusState.innerText = state.isExporting ? 'EXPORTING' : 'READY';
}

function updateTimeline() {
    // During drag, lock the timeline to the tile being scrubbed.
    const t = state.isDragging ? state.activeTile : (state.hoverTile || state.activeTile);
    const rangeDur = state.range.out - state.range.in;
    if (!t) {
        el.tlTile.innerText = 'No tile hovered';
        el.tlTime.innerText = `${formatTime(state.range.in)} / ${formatTime(state.range.out)}`;
        el.tlFill.style.width = '0%';
        el.tlHead.style.left = '0%';
        return;
    }
    const pct = (t.frameIndex / Math.max(1, state.totalFrames - 1)) * 100;
    const timeAt = state.range.in + (t.frameIndex / Math.max(1, state.totalFrames - 1)) * rangeDur;
    el.tlTile.innerText = `TILE ${String(t.id).padStart(2, '0')}${t.isPinned ? ' · PINNED' : ''} · FRAME ${t.frameIndex + 1} / ${state.totalFrames}`;
    el.tlTime.innerText = `${formatTime(timeAt)} · in ${formatTime(state.range.in)}–${formatTime(state.range.out)}`;
    el.tlFill.style.width = `${pct}%`;
    el.tlHead.style.left = `${pct}%`;
}

/* ==========================================================
   INTERACTION
   ========================================================== */

function tileAtEvent(e) {
    const r = canvas.getBoundingClientRect();
    const mx = (e.clientX - r.left) * (canvas.width / r.width);
    const my = (e.clientY - r.top) * (canvas.height / r.height);
    return state.tiles.find(t => mx >= t.x && mx <= t.x + t.w && my >= t.y && my <= t.y + t.h);
}

// Disable native touch gestures on the canvas so drag-to-scrub
// doesn't trigger scroll/pinch on touch devices.
canvas.style.touchAction = 'none';

canvas.addEventListener('pointerdown', (e) => {
    // Left mouse button, primary touch, or pen — ignore middle/right
    if (e.pointerType === 'mouse' && e.button !== 0) return;
    // Allow scrub in Standard mode even while playing (live time-scrubbing)
    if (anim.mode !== 'standard') return;
    const t = tileAtEvent(e);
    if (t && !t.isPinned) {
        state.activeTile = t;
        state.hoverTile = t;
        state.isDragging = true;
        state.scrubbingTile = t;
        state.dragStartX = e.clientX;
        state.dragStartFrame = t.frameOffset;
        document.body.style.cursor = 'ew-resize';
        // Capture so we keep receiving events if the pointer leaves the canvas
        try { canvas.setPointerCapture(e.pointerId); } catch (_) { /* no-op */ }
        renderAll();
        updateTimeline();
    }
});

canvas.addEventListener('pointermove', (e) => {
    // While dragging, don't let hover steal the timeline
    if (state.isDragging) {
        const r = canvas.getBoundingClientRect();
        const px = e.clientX - state.dragStartX;
        const sensitivity = Math.max(4, r.width / state.totalFrames * 1.2);
        const N = state.totalFrames;
        let newOffset = state.dragStartFrame + Math.floor(px / sensitivity);
        newOffset = Math.max(-(N - 1), Math.min(N - 1, newOffset));
        if (state.activeTile) {
            state.activeTile.frameOffset = newOffset;
            const rangeDur = Math.max(0.1, state.range.out - state.range.in);
            const rate = N / rangeDur;
            const outputFrame = anim.elapsed * rate;
            const stutterFrame = Math.floor(outputFrame / anim.stutter) * anim.stutter;
            const idx = applyLoop(Math.floor(stutterFrame + newOffset), N);
            if (state.activeTile.frameIndex !== idx) {
                state.activeTile.frameIndex = idx;
                renderAll();
                updateTimeline();
            }
        }
        return;
    }
    state.hoverTile = tileAtEvent(e);
    updateTimeline();
});

canvas.addEventListener('pointerleave', () => {
    if (state.isDragging) return;
    state.hoverTile = null;
    updateTimeline();
});

const endDrag = () => {
    if (state.isDragging) {
        state.hoverTile = state.activeTile;
        state.isDragging = false;
        state.scrubbingTile = null;
        state.activeTile = null;
        document.body.style.cursor = '';
        renderAll();
        updateTimeline();
    }
};
canvas.addEventListener('pointerup', endDrag);
canvas.addEventListener('pointercancel', endDrag);

// Right-click to pin (no touch equivalent yet — see long-press below)
canvas.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const t = tileAtEvent(e);
    if (t) {
        t.isPinned = !t.isPinned;
        renderAll();
        updateStatus();
        updateTimeline();
    }
});

// Long-press to pin on touch devices (500ms with <10px movement).
// On touch, pointerdown above always starts a scrub — if the user holds still,
// we treat it as a pin gesture instead and cancel the scrub.
let longPressTimer = null;
let longPressTile = null;
let longPressStartX = 0;

canvas.addEventListener('pointerdown', (e) => {
    if (e.pointerType !== 'touch') return;
    longPressTile = tileAtEvent(e);
    longPressStartX = e.clientX;
    if (!longPressTile) return;
    longPressTimer = setTimeout(() => {
        if (!longPressTile) return;
        // Toggle pin and cancel any in-flight scrub
        longPressTile.isPinned = !longPressTile.isPinned;
        state.isDragging = false;
        state.activeTile = null;
        state.hoverTile = longPressTile;
        if (navigator.vibrate) navigator.vibrate(30);
        renderAll();
        updateStatus();
        updateTimeline();
        longPressTile = null;
        longPressTimer = null;
    }, 500);
});

const clearLongPress = () => {
    if (longPressTimer) { clearTimeout(longPressTimer); longPressTimer = null; }
    longPressTile = null;
};
canvas.addEventListener('pointerup', clearLongPress);
canvas.addEventListener('pointercancel', clearLongPress);
canvas.addEventListener('pointermove', (e) => {
    // Any real movement (>10px) = scrub intent, not a hold
    if (longPressTimer && Math.abs(e.clientX - longPressStartX) > 10) {
        clearLongPress();
    }
});

/* ==========================================================
   NOISE UTILITIES
   ========================================================== */

function _seededRng(seed) {
    let s = seed | 0;
    return () => {
        s = (Math.imul(s, 1664525) + 1013904223) | 0;
        return (s >>> 0) / 4294967296;
    };
}

function _sstep(t) { return t * t * (3 - 2 * t); }

function _ihash3(x, y, z) {
    let h = ((x * 1619 + y * 31337 + z * 52711 + 1013904223) | 0);
    h = (Math.imul(h ^ (h >>> 13), 1664525) + 1013904223) | 0;
    return ((h ^ (h >>> 16)) >>> 0) / 4294967296;
}

function valueNoise3(x, y, z) {
    const xi = Math.floor(x), yi = Math.floor(y), zi = Math.floor(z);
    const fx = _sstep(x - xi), fy = _sstep(y - yi), fz = _sstep(z - zi);
    const L = (a, b, t) => a + (b - a) * t;
    const v = (dx, dy, dz) => _ihash3(xi + dx, yi + dy, zi + dz);
    return L(
        L(L(v(0,0,0), v(1,0,0), fx), L(v(0,1,0), v(1,1,0), fx), fy),
        L(L(v(0,0,1), v(1,0,1), fx), L(v(0,1,1), v(1,1,1), fx), fy),
        fz
    );
}

/* ==========================================================
   ANIMATION ENGINE
   ========================================================== */

function setMode(mode) {
    if (anim.raf) { cancelAnimationFrame(anim.raf); anim.raf = null; }
    anim.playing = false;
    anim.mode = mode;

    if (mode === 'standard') {
        canvas.style.cursor = 'ew-resize';
        anim.elapsed = 0;
        if (state.ready) renderAll();
        updatePlayPauseUI();
        return;
    }

    canvas.style.cursor = 'default';

    if (!state.ready) { updatePlayPauseUI(); return; }

    const n = state.tiles.length;
    const N = state.totalFrames;
    const rng = _seededRng(Date.now() | 0);
    anim.tileOffsets = Array.from({length: n}, () => rng());
    anim.tilePhases = anim.tileOffsets.map(r => r * N);
    anim.elapsed = 0;

    play();  // auto-start when picking an animated mode
}

function play() {
    if (!state.ready) return;
    if (anim.raf) return;
    // In hold mode, if we've already reached the end, restart from the beginning
    // so that pressing Play again replays instead of freezing immediately.
    if (anim.loopMode === 'hold') {
        const N = state.totalFrames;
        const rangeDur = Math.max(0.1, state.range.out - state.range.in);
        if (anim.elapsed * (N / rangeDur) >= N - 1) anim.elapsed = 0;
    }
    anim.lastNow = performance.now();
    const loop = (now) => {
        const dt = Math.min((now - anim.lastNow) / 1000, 0.1);
        anim.lastNow = now;
        anim.elapsed += dt;
        tickAnim(dt);
        renderAll();
        anim.raf = requestAnimationFrame(loop);
    };
    anim._loop = loop;
    anim.raf = requestAnimationFrame(loop);
    anim.playing = true;
    updatePlayPauseUI();
}

function pause() {
    if (anim.raf) { cancelAnimationFrame(anim.raf); anim.raf = null; }
    anim.playing = false;
    updatePlayPauseUI();
}

function togglePlayPause() {
    if (anim.playing) pause();
    else play();
}

function updatePlayPauseUI() {
    el.btnPlayPause.disabled = !state.ready;
    if (anim.playing) {
        el.playpauseIcon.innerText = '⏸';
        el.playpauseLabel.innerText = 'Pause';
        el.btnPlayPause.setAttribute('aria-label', 'Pause');
    } else {
        el.playpauseIcon.innerText = '▶';
        el.playpauseLabel.innerText = 'Play';
        el.btnPlayPause.setAttribute('aria-label', 'Play');
    }
}

function resetSpatialShuffle() {
    state.tiles.forEach(t => {
        t.srcX = t.origSrcX;
        t.srcY = t.origSrcY;
    });
}

function tickAnim(dt) {
    const cols = parseInt(el.inputCols.value);
    const N = state.totalFrames;
    const rangeDur = Math.max(0.1, state.range.out - state.range.in);
    const rate = N / rangeDur;
    const t = anim.elapsed;

    // Compute the output frame number (continuous, not quantized)
    const outputFrame = t * rate;

    state.tiles.forEach((tile, i) => {
        // Skip the tile currently being scrubbed by the user
        if (tile === state.scrubbingTile) return;

        // Use the pipeline: Pin → Stutter → Mode → Loop
        tile.frameIndex = computeFrameIndex(
            tile, outputFrame, cols, N, rate, t,
            anim.tileOffsets, anim.tilePhases, i, dt
        );
    });
    // applyLoop('hold') clamps each tile to N-1 — no separate pause needed here.
}

/* ==========================================================
   SPATIAL SHUFFLE
   ========================================================== */

function applySpatialShuffle(amount) {
    // Spatial Shuffle: swap the source crop positions (srcX/srcY) between tiles,
    // rearranging which region of the video each tile displays.
    // amount (0-10): 0 = no shuffle, 10 = full shuffle.
    const n = state.tiles.length;
    if (n < 2) return;

    // Work from original positions so repeated calls are idempotent
    const positions = state.tiles.map(t => ({ srcX: t.origSrcX, srcY: t.origSrcY }));

    const swapCount = Math.round((amount / 10) * n);

    for (let s = 0; s < swapCount; s++) {
        const a = Math.floor(Math.random() * n);
        const b = Math.floor(Math.random() * n);
        if (a !== b) {
            const tmp = positions[a];
            positions[a] = positions[b];
            positions[b] = tmp;
        }
    }

    state.tiles.forEach((t, i) => {
        t.srcX = positions[i].srcX;
        t.srcY = positions[i].srcY;
    });

    if (!anim.playing) renderAll();
}

/* ==========================================================
   BLOCK PATTERN
   ========================================================== */

function applyBlockPattern(pattern, amount) {
    const cols = parseInt(el.inputCols.value);
    const rows = parseInt(el.inputRows.value);
    const N = state.totalFrames;

    state.tiles.forEach((tile, i) => {
        const col = i % cols;
        const row = Math.floor(i / cols);
        let shouldPin = false;

        switch (pattern) {
            case 'checkerboard':
                shouldPin = (col + row) % 2 === 0;
                break;
            case 'random':
                shouldPin = Math.random() * 100 < amount;
                break;
            case 'borders':
                shouldPin = col === 0 || col === cols - 1 || row === 0 || row === rows - 1;
                break;
            case 'center':
                shouldPin = col > 0 && col < cols - 1 && row > 0 && row < rows - 1;
                break;
            default:
                shouldPin = false;
        }

        tile.isPinned = shouldPin;
        if (shouldPin) {
            tile.frameIndex = Math.floor(Math.random() * N);
        }
    });

    renderAll();
    updateStatus();
    updateTimeline();
}

function clearAllPins() {
    state.tiles.forEach(t => {
        t.isPinned = false;
        t.frameIndex = 0;
    });
    renderAll();
    updateStatus();
    updateTimeline();
}

/* ==========================================================
   LOADING OVERLAY
   ========================================================== */

function showLoading(msg) {
    if (el.loadingText) el.loadingText.innerText = msg;
    el.loadingOverlay.classList.add('visible');
    el.progressFill.style.width = '0%';
}

function hideLoading() {
    el.loadingOverlay.classList.remove('visible');
}

/* ==========================================================
   CONTROLS
   ========================================================== */

function enableControls() {
    el.inputCols.disabled = false;
    el.inputRows.disabled = false;
    el.checkSquare.disabled = false;
    el.checkGrid.disabled = false;
    el.btnExportPng.disabled = false;
    el.btnExportVideo.disabled = false;
    el.selectMode.disabled = false;
    el.inputSpatialAmt.disabled = false;
    el.checkSpatialShuffle.disabled = false;
    el.selectPattern.disabled = false;
    el.inputPatternAmt.disabled = false;
    el.btnPatternApply.disabled = false;
    el.btnPatternClear.disabled = false;
    el.inputDuration.disabled = false;
    el.selectRes.disabled = false;
    el.selectFps.disabled = false;
    el.checkLoop.disabled = false;
    el.btnPlayPause.disabled = false;
    if (el.selectLoop) el.selectLoop.disabled = false;
    if (el.inputStutter) el.inputStutter.disabled = false;
}

/* ==========================================================
   EXPORT ENGINE — Dual Path
   ========================================================== */

/**
 * Automatically choose the best export path.
 * Path A: WebCodecs + mp4-muxer (Primary - MP4 H.264)
 *   Target: Chrome, Edge, Safari 16.4+, Opera
 * Path B: MediaRecorder (Fallback - WebM)
 *   Target: Firefox, old Safari
 */
function getExportPath() {
    const hasWebCodecs = typeof VideoEncoder !== 'undefined' && typeof VideoFrame !== 'undefined';
    const hasMp4Muxer = typeof Mp4Muxer !== 'undefined';
    if (hasWebCodecs && hasMp4Muxer) return 'A';
    return 'B';
}

/**
 * Ensure canvas dimensions are even (required for H.264 compliance).
 */
function ensureEvenDimensions(w, h) {
    return {
        width: w + (w % 2),
        height: h + (h % 2)
    };
}

/**
 * Export video — always stops preview first.
 */
async function exportVideo() {
    if (state.isExporting || !state.ready || extractionRunning) return;

    // Stop preview before export
    pause();

    state.isExporting = true;
    updateStatus();
    const path = getExportPath();
    // Path B records in real time and stalls in background tabs
    showLoading(path === 'A' ? 'EXPORTING...' : 'EXPORTING IN REAL TIME — KEEP THIS TAB VISIBLE');

    console.log(`Export path: ${path} (${path === 'A' ? 'WebCodecs+mp4-muxer' : 'MediaRecorder fallback'})`);

    try {
        if (path === 'A') {
            await exportVideoWebCodecs();
        } else {
            await exportVideoMediaRecorder();
        }
    } catch (err) {
        console.error('Export failed:', err);
        showToast(`Export failed: ${err.message || 'Unknown error'}`);
    } finally {
        hideLoading();
        state.isExporting = false;
        updateStatus();
    }
}

/**
 * Path A: WebCodecs + mp4-muxer (Primary - MP4 H.264)
 */
async function exportVideoWebCodecs() {
    const fps = parseInt(el.selectFps.value, 10);
    const dur = parseFloat(el.inputDuration.value);
    const res = el.selectRes.value;
    const totalFrames = Math.ceil(dur * fps);

    // Determine export canvas size
    let expW, expH;
    if (res === 'preview') {
        expW = canvas.width;
        expH = canvas.height;
    } else {
        // Scale the canvas aspect ratio to fit the chosen resolution tier.
        // Hardcoded 16:9 dims would stretch portrait or square content.
        const longSide = { '720': 1280, '1080': 1920, '4k': 3840 }[res] || 1920;
        const aspect = canvas.width / canvas.height;
        if (aspect >= 1) { expW = longSide; expH = Math.round(longSide / aspect); }
        else             { expH = longSide; expW = Math.round(longSide * aspect); }
    }
    const even = ensureEvenDimensions(expW, expH);
    expW = even.width;
    expH = even.height;

    // Bitrate by resolution
    const bitrateMap = { 'preview': 6000000, '720': 6000000, '1080': 10000000, '4k': 20000000 };
    let bitrate = bitrateMap[res] || 10000000;

    // Create export canvas
    const expCanvas = document.createElement('canvas');
    expCanvas.width = expW;
    expCanvas.height = expH;
    const expCtx = expCanvas.getContext('2d');
    // Fresh phase state so the export is deterministic and never mutates the preview
    const exportPhases = createInitialPhases();

    // Setup mp4-muxer
    const muxer = new Mp4Muxer.Muxer({
        target: new Mp4Muxer.ArrayBufferTarget(),
        video: {
            codec: 'avc',
            width: expW,
            height: expH,
            bitrate: bitrate
        },
        fastStart: 'in-memory'
    });

    // Pick the minimum H.264 level that covers the actual pixel area.
    // Using a level too low throws "coded area exceeds maximum" even after the
    // aspect-ratio fix (e.g. 732×1280 = 937k px > L3.1 limit of 921k px).
    const codecString = (() => {
        const px = expW * expH;
        if (px <= 921600)  return 'avc1.42001f'; // Baseline L3.1 ≤ 1280×720
        if (px <= 2097152) return 'avc1.420028'; // Baseline L4.0 ≤ ~1920×1080
        if (px <= 9437184) return 'avc1.420033'; // Baseline L5.1 ≤ 3840×2160
        return 'avc1.420034';                     // Baseline L5.2 for anything larger
    })();

    // Capture encoder errors via variable — throwing inside WebCodecs callbacks
    // does NOT propagate to the outer async try/catch; it only closes the encoder.
    let encoderError = null;
    let frameCount = 0;
    const encoder = new VideoEncoder({
        output: (chunk, meta) => muxer.addVideoChunk(chunk, meta),
        error: (err) => { encoderError = err; }
    });

    const encoderConfig = {
        codec: codecString,
        width: expW,
        height: expH,
        bitrate: bitrate,
        framerate: fps
    };

    // Fail fast (before rendering anything) if the GPU/browser can't encode this size
    const support = await VideoEncoder.isConfigSupported(encoderConfig).catch(() => null);
    if (!support || !support.supported) {
        encoder.close();
        throw new Error(`H.264 ${expW}×${expH} is not supported by this browser — try a lower resolution`);
    }
    encoder.configure(encoderConfig);

    // Render each frame
    for (let i = 0; i < totalFrames; i++) {
        // Backpressure: wait if queue is too large
        while (encoder.encodeQueueSize > 10) {
            await new Promise(r => setTimeout(r, 10));
        }

        // Re-throw the real encoder error if one occurred
        if (encoderError) throw encoderError;
        if (encoder.state === 'closed') {
            throw new Error('VideoEncoder closed unexpectedly during export');
        }

        // Render the frame at output frame i
        renderExportFrame(expCtx, expCanvas, i, totalFrames, fps, exportPhases);

        // Create VideoFrame
        const videoFrame = new VideoFrame(expCanvas, {
            timestamp: i * 1_000_000 / fps, // microseconds
            duration: Math.round(1_000_000 / fps)
        });

        encoder.encode(videoFrame);
        videoFrame.close(); // GC: close immediately after encode

        frameCount++;

        // Update progress
        el.progressFill.style.width = `${((i + 1) / totalFrames) * 100}%`;
    }

    // Wait for all pending encodes to complete before flushing
    // This prevents "Cannot call 'encode' on a closed codec" errors
    while (encoder.encodeQueueSize > 0) {
        await new Promise(r => setTimeout(r, 10));
    }

    // Finalize
    await encoder.flush();
    encoder.close();
    muxer.finalize();

    // Get the buffer
    const buffer = muxer.target.buffer;
    const blob = new Blob([buffer], { type: 'video/mp4' });

    // Generate filename
    const resLabel = res === 'preview' ? `${canvas.width}x${canvas.height}` : res;
    const ts = Date.now();
    const filename = `PanoTile-${resLabel}-${fps}fps-${dur}s-${ts}.mp4`;

    // Download
    downloadBlob(blob, filename);

    showToast(`Exported ${filename} (${formatBytes(blob.size)})`, 'info');
}

/**
 * Render a single export frame at the given output frame index.
 * Uses the pipeline to compute tile frame indices.
 */
function renderExportFrame(expCtx, expCanvas, outputFrame, totalFrames, fps, exportPhases) {
    expCtx.clearRect(0, 0, expCanvas.width, expCanvas.height);

    const cols = parseInt(el.inputCols.value);
    const N = state.totalFrames;
    const rangeDur = Math.max(0.1, state.range.out - state.range.in);
    const rate = N / rangeDur;
    const elapsed = outputFrame / fps; // time in seconds at this output frame
    // Same time base as the preview (tickAnim): source frames, not output frames.
    // They differ whenever the source frame count was clamped or fps changed.
    const sourceFrame = elapsed * rate;

    // Scale tile positions from preview canvas to export canvas
    const scaleX = expCanvas.width / canvas.width;
    const scaleY = expCanvas.height / canvas.height;

    for (let i = 0; i < state.tiles.length; i++) {
        const tile = state.tiles[i];

        // Compute frame index via pipeline
        const fi = computeFrameIndex(
            tile, sourceFrame, cols, N, rate, elapsed,
            anim.tileOffsets, exportPhases, i, 1 / fps
        );

        const img = state.frames[fi];
        if (img) {
            // Draw at export resolution
            expCtx.drawImage(
                img,
                tile.srcX, tile.srcY, tile.srcW, tile.srcH,
                tile.x * scaleX, tile.y * scaleY,
                tile.w * scaleX, tile.h * scaleY
            );
        }
    }
}

/**
 * Path B: MediaRecorder (Fallback - WebM)
 */
async function exportVideoMediaRecorder() {
    const fps = parseInt(el.selectFps.value, 10);
    const dur = parseFloat(el.inputDuration.value);
    const res = el.selectRes.value;
    const totalFrames = Math.ceil(dur * fps);

    // Determine export canvas size
    let expW, expH;
    if (res === 'preview') {
        expW = canvas.width;
        expH = canvas.height;
    } else {
        // Scale the canvas aspect ratio to fit the chosen resolution tier.
        // Hardcoded 16:9 dims would stretch portrait or square content.
        const longSide = { '720': 1280, '1080': 1920, '4k': 3840 }[res] || 1920;
        const aspect = canvas.width / canvas.height;
        if (aspect >= 1) { expW = longSide; expH = Math.round(longSide / aspect); }
        else             { expH = longSide; expW = Math.round(longSide * aspect); }
    }
    const even = ensureEvenDimensions(expW, expH);
    expW = even.width;
    expH = even.height;

    // Create export canvas
    const expCanvas = document.createElement('canvas');
    expCanvas.width = expW;
    expCanvas.height = expH;
    const expCtx = expCanvas.getContext('2d');
    // Fresh phase state so the export is deterministic and never mutates the preview
    const exportPhases = createInitialPhases();

    // Setup MediaRecorder — captureStream(0) is manual (frames pushed via requestFrame).
    // If requestFrame is unavailable, fall back to a timed capture at the export fps.
    let stream = expCanvas.captureStream(0);
    let videoTrack = stream.getVideoTracks()[0];
    if (!videoTrack || typeof videoTrack.requestFrame !== 'function') {
        for (const track of stream.getTracks()) track.stop();
        stream = expCanvas.captureStream(fps);
        videoTrack = stream.getVideoTracks()[0];
    }

    // Prefer VP9, fallback to VP8
    const mimeType = ['video/webm;codecs=vp9', 'video/webm;codecs=vp8', 'video/webm']
        .find(type => MediaRecorder.isTypeSupported(type)) || '';

    const recorder = new MediaRecorder(stream, mimeType ? { mimeType } : undefined);
    const chunks = [];

    recorder.ondataavailable = (e) => {
        if (e.data.size > 0) chunks.push(e.data);
    };

    recorder.start(100); // timeslice 100ms

    // MediaRecorder timestamps frames with wall-clock time, so frames must be
    // pushed exactly 1/fps apart — pacing on rAF (display refresh) would make
    // a 30fps export play at 2× on a 60Hz screen. This path runs in real time.
    const frameMs = 1000 / fps;
    const t0 = performance.now();
    for (let i = 0; i < totalFrames; i++) {
        renderExportFrame(expCtx, expCanvas, i, totalFrames, fps, exportPhases);
        if (videoTrack && typeof videoTrack.requestFrame === 'function') {
            videoTrack.requestFrame();
        }

        el.progressFill.style.width = `${((i + 1) / totalFrames) * 100}%`;

        const wait = t0 + (i + 1) * frameMs - performance.now();
        await new Promise(r => setTimeout(r, Math.max(0, wait)));
    }

    // Flush: wait 300ms before stopping
    await new Promise(r => setTimeout(r, 300));

    // Set onstop BEFORE calling stop() to avoid race condition
    const stopPromise = new Promise(r => { recorder.onstop = r; });
    recorder.stop();
    await stopPromise;
    for (const track of stream.getTracks()) track.stop();

    const blob = new Blob(chunks, { type: mimeType || 'video/webm' });

    // Generate filename
    const resLabel = res === 'preview' ? `${canvas.width}x${canvas.height}` : res;
    const ts = Date.now();
    const filename = `PanoTile-${resLabel}-${fps}fps-${dur}s-${ts}.webm`;

    downloadBlob(blob, filename);

    showToast(`Exported ${filename} (${formatBytes(blob.size)})`, 'info');
}

/**
 * Phase state at elapsed = 0 — same initialisation setMode() uses for the preview.
 */
function createInitialPhases() {
    return anim.tileOffsets.map(r => r * state.totalFrames);
}

/**
 * Download a blob as a file.
 */
function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    // Revoke blob URL after 2 seconds (GC)
    setTimeout(() => URL.revokeObjectURL(url), 2000);
}

/**
 * Export PNG (single frame snapshot).
 */
function exportPng() {
    // Stop preview before export
    pause();

    const res = el.selectRes.value;
    let expW, expH;
    if (res === 'preview') {
        expW = canvas.width;
        expH = canvas.height;
    } else {
        // Scale the canvas aspect ratio to fit the chosen resolution tier.
        // Hardcoded 16:9 dims would stretch portrait or square content.
        const longSide = { '720': 1280, '1080': 1920, '4k': 3840 }[res] || 1920;
        const aspect = canvas.width / canvas.height;
        if (aspect >= 1) { expW = longSide; expH = Math.round(longSide / aspect); }
        else             { expH = longSide; expW = Math.round(longSide * aspect); }
    }

    const expCanvas = document.createElement('canvas');
    expCanvas.width = expW;
    expCanvas.height = expH;
    const expCtx = expCanvas.getContext('2d');

    const scaleX = expW / canvas.width;
    const scaleY = expH / canvas.height;

    for (const tile of state.tiles) {
        const img = state.frames[tile.frameIndex];
        if (img) {
            expCtx.drawImage(
                img,
                tile.srcX, tile.srcY, tile.srcW, tile.srcH,
                tile.x * scaleX, tile.y * scaleY,
                tile.w * scaleX, tile.h * scaleY
            );
        }
    }

    const link = document.createElement('a');
    link.download = `PanoTile-${res}-${Date.now()}.png`;
    link.href = expCanvas.toDataURL('image/png');
    link.click();
}

/* ==========================================================
   EVENT BINDINGS
   ========================================================== */

// Mode selector
el.selectMode.addEventListener('change', () => {
    setMode(el.selectMode.value);
});

// Play/Pause
el.btnPlayPause.addEventListener('click', togglePlayPause);

// Spatial shuffle toggle
el.checkSpatialShuffle.addEventListener('change', () => {
    if (el.checkSpatialShuffle.checked) {
        applySpatialShuffle(parseInt(el.inputSpatialAmt.value));
    } else {
        resetSpatialShuffle();
    }
    if (!anim.playing) renderAll();
});

el.inputSpatialAmt.addEventListener('input', () => {
    el.spatialAmtValue.innerText = el.inputSpatialAmt.value;
    if (el.checkSpatialShuffle.checked) {
        resetSpatialShuffle();
        applySpatialShuffle(parseInt(el.inputSpatialAmt.value));
        if (!anim.playing) renderAll();
    }
});

// Pattern
el.selectPattern.addEventListener('change', () => {
    el.patternAmtRow.style.display = el.selectPattern.value === 'random' ? 'flex' : 'none';
});

el.btnPatternApply.addEventListener('click', () => {
    const pattern = el.selectPattern.value;
    if (pattern === 'none') return;
    const amount = parseInt(el.inputPatternAmt.value);
    applyBlockPattern(pattern, amount);
});

el.btnPatternClear.addEventListener('click', clearAllPins);

// Generate
// Export
el.btnExportVideo.addEventListener('click', exportVideo);
el.btnExportPng.addEventListener('click', exportPng);

// Duration slider
el.inputDuration.addEventListener('input', () => {
    el.durationValue.innerText = parseFloat(el.inputDuration.value).toFixed(1);
});

// Resolution / FPS
// FPS defines how many source frames are sampled from the range — re-extract on change
el.selectFps.addEventListener('change', () => {
    if (state.ready) reextractFrames();
});

// Loop checkbox → maps to loop mode dropdown
el.checkLoop.addEventListener('change', () => {
    // When loop checkbox is toggled, sync to loop mode
    if (el.checkLoop.checked) {
        anim.loopMode = 'wrap';
        if (el.selectLoop) el.selectLoop.value = 'wrap';
    } else {
        anim.loopMode = 'hold';
        if (el.selectLoop) el.selectLoop.value = 'hold';
    }
});

// Loop mode dropdown
if (el.selectLoop) {
    el.selectLoop.addEventListener('change', () => {
        anim.loopMode = el.selectLoop.value;
        // Sync checkbox
        el.checkLoop.checked = anim.loopMode === 'wrap';
    });
}

// Stutter slider
if (el.inputStutter) {
    el.inputStutter.addEventListener('input', () => {
        anim.stutter = parseInt(el.inputStutter.value, 10);
        if (el.stutterValue) el.stutterValue.innerText = anim.stutter;
    });
}

// Grid toggle
el.checkGrid.addEventListener('change', () => {
    if (!anim.playing) renderAll();
});

// Grid changes rebuild every tile — ask first if the user has pinned or scrubbed tiles.
const gridApplied = { cols: el.inputCols.value, rows: el.inputRows.value, square: el.checkSquare.checked };

function rememberGrid() {
    gridApplied.cols = el.inputCols.value;
    gridApplied.rows = el.inputRows.value;
    gridApplied.square = el.checkSquare.checked;
}

function hasTileEdits() {
    return state.tiles.some(t => t.isPinned || t.frameOffset !== 0);
}

async function rebuildGrid() {
    if (!state.ready) return;
    const unchanged = el.inputCols.value === gridApplied.cols
        && el.inputRows.value === gridApplied.rows
        && el.checkSquare.checked === gridApplied.square;
    if (unchanged) return;
    if (hasTileEdits()) {
        const ok = await confirm(
            'Rebuild grid?',
            'Changing the grid resets all pinned tiles and scrub offsets.',
            'Rebuild'
        );
        if (!ok) {
            el.inputCols.value = gridApplied.cols;
            el.inputRows.value = gridApplied.rows;
            el.checkSquare.checked = gridApplied.square;
            el.inputRows.disabled = gridApplied.square;
            return;
        }
    }
    pause();
    initProject();
    renderAll();
}

el.checkSquare.addEventListener('change', rebuildGrid);
el.inputCols.addEventListener('change', rebuildGrid);
el.inputRows.addEventListener('change', () => {
    if (!el.checkSquare.checked) rebuildGrid();
});

/* ==========================================================
   KEYBOARD SHORTCUTS
   ========================================================== */

document.addEventListener('keydown', (e) => {
    // Ignore if user is typing in an input
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;
    // Leave browser/OS shortcuts alone (Cmd+S, Ctrl+R, …)
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    // No shortcuts behind a modal, while loading or while exporting
    if ($('tutorialOverlay')?.classList.contains('visible')) return;
    if (el.confirmDialog.classList.contains('visible')) return;
    if (el.loadingOverlay.classList.contains('visible') || state.isExporting) return;
    // Space on a focused button should only activate that button
    if (e.key === ' ' && e.target.closest?.('button')) return;

    switch (e.key.toLowerCase()) {
        case ' ': // Space: Play/Pause
            e.preventDefault();
            togglePlayPause();
            break;
        case 'l': // Linear L->R
            el.selectMode.value = 'linear-lr';
            setMode('linear-lr');
            break;
        case 'k': // Linear R->L
            el.selectMode.value = 'linear-rl';
            setMode('linear-rl');
            break;
        case 's': // Shuffle
            el.selectMode.value = 'temporal-shuffle';
            setMode('temporal-shuffle');
            break;
        case 'p': // Perlin flow
            el.selectMode.value = 'perlin-flow';
            setMode('perlin-flow');
            break;
        case 'd': // Drunk walk
            el.selectMode.value = 'drunk';
            setMode('drunk');
            break;
        case 'r': // Reset time (pause + reset elapsed)
            pause();
            anim.elapsed = 0;
            if (anim.mode === 'standard') {
                state.tiles.forEach(t => { if (!t.isPinned) t.frameIndex = 0; });
            }
            renderAll();
            break;
        case 'g': // Toggle grid
            el.checkGrid.checked = !el.checkGrid.checked;
            if (!anim.playing) renderAll();
            break;
    }
});

/* ==========================================================
   TUTORIAL MODAL
   ========================================================== */

let tutorialStep = 0;
const TUTORIAL_STEPS = 4;

function openTutorial() {
    const overlay = $('tutorialOverlay');
    if (!overlay) return;
    tutorialStep = 0;
    showTutorialStep(0);
    overlay.classList.add('visible');
}

function closeTutorial() {
    const overlay = $('tutorialOverlay');
    if (!overlay) return;
    overlay.classList.remove('visible');
}

function showTutorialStep(index) {
    tutorialStep = Math.max(0, Math.min(TUTORIAL_STEPS - 1, index));

    // Update steps visibility
    document.querySelectorAll('.tutorial-step').forEach(el => {
        el.classList.toggle('active', parseInt(el.dataset.step) === tutorialStep);
    });

    // Update dots
    document.querySelectorAll('.tutorial-dot').forEach(dot => {
        dot.classList.toggle('active', parseInt(dot.dataset.index) === tutorialStep);
    });

    // Update nav buttons
    const prevBtn = $('tutorialPrev');
    const nextBtn = $('tutorialNext');
    if (prevBtn) prevBtn.style.visibility = tutorialStep === 0 ? 'hidden' : 'visible';
    if (nextBtn) {
        if (tutorialStep === TUTORIAL_STEPS - 1) {
            nextBtn.textContent = 'Got it!';
        } else {
            nextBtn.textContent = 'Next →';
        }
    }
}

function goToPrevStep() {
    showTutorialStep(tutorialStep - 1);
}

function goToNextStep() {
    if (tutorialStep === TUTORIAL_STEPS - 1) {
        closeTutorial();
    } else {
        showTutorialStep(tutorialStep + 1);
    }
}

/* ==========================================================
   INIT
   ========================================================== */

// Initial state
updatePlayPauseUI();

// Tutorial: auto-show on first visit
// Storage can throw (private mode, blocked site data) — treat as first visit, never crash init
let tutorialSeen = false;
try { tutorialSeen = !!localStorage.getItem('panotile_tutorial_seen'); } catch (_) {}
if (!tutorialSeen) {
    try { localStorage.setItem('panotile_tutorial_seen', '1'); } catch (_) {}
    // Small delay to let DOM settle
    setTimeout(openTutorial, 300);
}

// Tutorial event bindings
const btnHelp = $('btnHelp');
if (btnHelp) btnHelp.addEventListener('click', openTutorial);

const tutorialOverlay = $('tutorialOverlay');
if (tutorialOverlay) {
    tutorialOverlay.addEventListener('click', (e) => {
        if (e.target === tutorialOverlay) closeTutorial();
    });
}

const tutorialClose = $('tutorialClose');
if (tutorialClose) tutorialClose.addEventListener('click', closeTutorial);

const tutorialPrev = $('tutorialPrev');
if (tutorialPrev) tutorialPrev.addEventListener('click', goToPrevStep);

const tutorialNext = $('tutorialNext');
if (tutorialNext) tutorialNext.addEventListener('click', goToNextStep);

// Dot navigation
document.querySelectorAll('.tutorial-dot').forEach(dot => {
    dot.addEventListener('click', () => {
        showTutorialStep(parseInt(dot.dataset.index));
    });
});

