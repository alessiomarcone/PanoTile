'use strict';

/* ==========================================================
   STATE
   ========================================================== */

// Source frame count is now computed from range duration × target fps.
// These are guardrails to keep memory sane.
const MIN_SOURCE_FRAMES = 30;     // below this, scrubbing feels choppy
const MAX_SOURCE_FRAMES = 1800;   // 60s × 30fps — hard ceiling

// The video is streamed from a blob URL (never read into memory), so the file
// size limit is generous — memory is governed by the frame cache budget below.
const MAX_FILE_BYTES = 4 * 1024 * 1024 * 1024;
const MAX_CANVAS_DIM = 1000;

// Frame cache: decoded frames are stored as RGBA bitmaps (4 bytes/px).
// They are downscaled so the whole cache fits a memory budget.
const MIN_CACHE_LONG_SIDE = 480;
const MAX_EXPORT_DURATION = 300;

const state = {
    video: { width: 0, height: 0, duration: 0, name: '', size: 0, blobUrl: null, probe: null, file: null },
    // Frame cache size relative to the source video (tile src coords are in source pixels)
    frameScaleX: 1,
    frameScaleY: 1,
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
    selectCache: $('selectCache'),
    cacheInfo: $('cacheInfo'),
    exportQualityWarn: $('exportQualityWarn'),
    inputFreezeSpan: $('inputFreezeSpan'),
    freezeSpanValue: $('freezeSpanValue'),
    btnUnfreeze: $('btnUnfreeze'),
    btnProjectSave: $('btnProjectSave'),
    btnProjectOpen: $('btnProjectOpen'),
    projectUpload: $('projectUpload'),
    btnUndo: $('btnUndo'),
    btnRedo: $('btnRedo'),
    btnCancelTask: $('btnCancelTask'),
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
    if (b < 1024 * 1024 * 1024) return (b / (1024 * 1024)).toFixed(1) + ' MB';
    return (b / (1024 * 1024 * 1024)).toFixed(2) + ' GB';
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
        showToast(`File too large (${formatBytes(file.size)}, max ${formatBytes(MAX_FILE_BYTES)})`);
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
    lastExtracted = null;
    state.tiles = [];
    state.hoverTile = null;
    state.activeTile = null;
    state.video.name = file.name;
    state.video.size = file.size;
    state.video.blobUrl = URL.createObjectURL(file);
    state.video.file = file;   // kept for the WebCodecs fast path (reads byte ranges)

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
        el.inputDuration.max = MAX_EXPORT_DURATION;

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

/**
 * Memory budget for the frame cache. navigator.deviceMemory is Chromium-only;
 * elsewhere assume 4 GB. Budget = 15% of it, clamped 0.5–1.5 GB — a browser tab
 * also needs room for decoding, export canvases and the encoder.
 */
function frameCacheBudget() {
    const gb = navigator.deviceMemory || 4;
    return Math.min(1.5, Math.max(0.5, gb * 0.15)) * 1024 * 1024 * 1024;
}

/** Size of each cached frame for `count` frames, honouring the quality selector. */
function computeCacheSize(count) {
    const vW = state.video.width, vH = state.video.height;
    const longSide = Math.max(vW, vH);
    const pref = el.selectCache ? el.selectCache.value : 'auto';
    const budget = frameCacheBudget();
    let scale = 1;

    if (pref === 'auto') {
        // No bigger than the export needs, no bigger than the memory budget allows
        scale = Math.min(1, neededCacheScale(), Math.sqrt(budget / (count * vW * vH * 4)));
        scale = Math.max(scale, Math.min(1, MIN_CACHE_LONG_SIDE / longSide));
    } else if (pref !== 'full') {
        scale = Math.min(1, parseInt(pref, 10) / longSide);
    }

    // Auto rounds down so the result never lands a few bytes over the budget
    const round = pref === 'auto' ? Math.floor : Math.round;
    const w = Math.max(2, round(vW * scale));
    const h = Math.max(2, round(vH * scale));
    const bytes = count * w * h * 4;
    return { w, h, bytes, budget, overBudget: bytes > budget };
}

/** Source pixels shown across the canvas width (the grid crops the video to the canvas aspect). */
function sourceCropWidth(aspect) {
    const vW = state.video.width, vH = state.video.height;
    return aspect > vW / vH ? vW : vH * aspect;
}

function canvasAspect() {
    return state.tiles.length ? canvas.width / canvas.height : state.video.width / state.video.height;
}

/** Cache scale at which the selected export resolution shows cached pixels 1:1. */
function neededCacheScale() {
    const aspect = canvasAspect();
    const { w } = exportSize(el.selectRes.value, { aspect });
    return Math.min(1, w / sourceCropWidth(aspect));
}

/** Warn when the export has to enlarge the cached frames (soft / blocky output). */
function updateExportQualityWarning() {
    const warn = el.exportQualityWarn;
    if (!warn) return;
    if (!state.ready) { warn.hidden = true; return; }
    const aspect = canvasAspect();
    const { w } = exportSize(el.selectRes.value, { aspect });
    const cropW = sourceCropWidth(aspect);
    const up = w / (cropW * state.frameScaleX);
    if (up <= 1.15) { warn.hidden = true; return; }
    warn.hidden = false;
    warn.textContent = state.frameScaleX >= 0.999
        ? `ℹ The source is only ${Math.round(cropW)} px wide here — this export enlarges it ×${up.toFixed(1)}.`
        : `⚠ Frames are cached at ${Math.round(cropW * state.frameScaleX)} px — this export enlarges them ×${up.toFixed(1)} and will look soft. Raise Frame quality (Range), shorten the range or lower FPS.`;
}

/** Live estimate shown under the range selector. */
function updateCacheInfo() {
    if (!el.cacheInfo || !state.video.width) return;
    const count = computeSourceFrameCount();
    const c = computeCacheSize(count);
    el.cacheInfo.textContent = `${count} frames · ${c.w}×${c.h} · ≈${formatBytes(c.bytes)}`
        + (c.overBudget ? ` — over the ~${formatBytes(c.budget)} budget, may crash` : '');
    el.cacheInfo.classList.toggle('warn', c.overBudget);
}

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
// Range / fps / quality the current frames were sampled with
let lastExtracted = null;

function restoreExtractedSettings() {
    if (!lastExtracted) return;
    state.range = { ...lastExtracted.range };
    el.selectFps.value = lastExtracted.fps;
    el.selectCache.value = lastExtracted.cache;
    renderRangeUI();
}

/** Stop an in-flight extraction; a re-extraction falls back to the previous frames. */
function abortExtraction() {
    cancelExtraction();
    extractionRunning = false;
    hideLoading();
    state.ready = state.frames.length > 0 && state.tiles.length > 0;
    if (state.ready) {
        restoreExtractedSettings();
        renderAll();
    }
    showToast('Frame extraction cancelled', 'info');
}

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
    showLoading(`EXTRACTING ${targetCount} FRAMES`);
    setCancellable(abortExtraction);

    const cache = computeCacheSize(targetCount);
    const temp = document.createElement('canvas');
    temp.width = cache.w;
    temp.height = cache.h;
    const tCtx = temp.getContext('2d');
    tCtx.imageSmoothingQuality = 'high';

    const rangeDur = state.range.out - state.range.in;
    const targets = Array.from({ length: targetCount },
        (_, i) => state.range.in + (i / Math.max(1, targetCount - 1)) * rangeDur);
    const isCancelled = () => token !== extractionToken;
    const onProgress = (done) => {
        el.progressFill.style.width = `${(done / targetCount) * 100}%`;
    };

    // Fast path: demux + sequential WebCodecs decode (MP4/MOV). Falls back to
    // per-frame seeking for anything it can't handle.
    let frames = null;
    let method = 'seek';
    if (canUseFastExtraction(state.video.file)) {
        try {
            frames = await extractFramesWebCodecs(state.video.file, { targets, tCtx, temp, isCancelled, onProgress });
            method = 'webcodecs';
        } catch (err) {
            if (isCancelled()) return false;
            console.info(`Fast extraction unavailable (${err.message}) — using seek extraction`);
            frames = null;
        }
    }
    if (!frames) {
        el.progressFill.style.width = '0%';
        frames = await extractFramesSeek(probe, { targets, tCtx, temp, isCancelled, onProgress });
    }
    if (isCancelled()) {
        if (frames) closeBitmaps(frames);
        return false;
    }
    if (!frames) {
        extractionRunning = false;
        // A failed re-extraction keeps the previous frames (and their range) usable
        state.ready = state.frames.length > 0 && state.tiles.length > 0;
        if (state.ready) restoreExtractedSettings();
        hideLoading();
        showToast('Frame extraction failed. Try a different video.');
        return false;
    }
    window.__lastExtractMethod = method;   // read by the test harness
    console.info(`Extracted ${targetCount} frames via ${method}`);

    // Swap in the new frames only once extraction completed
    releaseFrames();
    state.frames = frames;
    state.totalFrames = targetCount;
    state.frameScaleX = cache.w / state.video.width;
    state.frameScaleY = cache.h / state.video.height;
    lastExtracted = { range: { ...state.range }, fps: el.selectFps.value, cache: el.selectCache.value };
    extractionRunning = false;

    hideLoading();
    canvas.style.display = 'block';
    el.emptyState.style.display = 'none';
    el.timeline.classList.add('visible');
    state.ready = true;
    enableControls();
    const fresh = !(keepProject && state.tiles.length > 0);
    if (fresh) {
        initProject();
    } else {
        remapProjectFrames(prevRange || { ...state.range }, prevN);
    }
    updateStatus();
    updateCacheInfo();
    updateExportQualityWarning();
    onFramesReady({ fresh });
    return true;
}

function closeBitmaps(frames) {
    for (const f of new Set(frames)) if (f) f.close();
}

/**
 * Original extraction: seek the <video> element to every target time.
 * Works for every format the browser plays, but each seek decodes from the
 * previous keyframe. Returns the frames, or null on failure/cancel.
 */
async function extractFramesSeek(probe, { targets, tCtx, temp, isCancelled, onProgress }) {
    const frames = [];
    for (let i = 0; i < targets.length; i++) {
        if (isCancelled()) { closeBitmaps(frames); return null; }
        try {
            frames.push(await seekAndCapture(probe, tCtx, temp, targets[i]));
        } catch (err) {
            if (isCancelled()) { closeBitmaps(frames); return null; }
            console.warn(`Frame ${i} failed:`, err);
            if (frames.length === 0) return null;
            frames.push(frames[frames.length - 1]);
        }
        onProgress(i + 1);
    }
    return frames;
}

/**
 * Keep pins/offsets after a re-extraction.
 * Pinned frames stay on the same moment of the video (clamped to the new range);
 * offsets and phases keep the same duration in seconds.
 */
function remapProjectFrames(prevRange, prevN) {
    const k = remapTileFrames(prevRange, prevN);
    anim.tilePhases = anim.tilePhases.map(p => p * k);
    renderAll();
    updateTimeline();
}

/**
 * Map tile frame indices/offsets sampled with (prevRange, prevN) onto the
 * current range and frame count. Returns the frame-rate ratio new/prev.
 */
function remapTileFrames(prevRange, prevN) {
    const newN = state.totalFrames;
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
    return k;
}

/** Re-extract frames after range or fps change. */
function reextractFrames() {
    if (!state.video.probe || state.isExporting) return;
    const hadProject = state.tiles.length > 0;
    pause();
    state.ready = false;
    extractFrames(state.video.probe, { keepProject: hadProject, prevRange: lastExtracted && lastExtracted.range });
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
                tCtx.drawImage(probe, 0, 0, temp.width, temp.height);
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
   FAST FRAME EXTRACTION — mp4box.js demux + WebCodecs decode
   ========================================================== */

// Seeking a <video> element costs a decode from the previous keyframe for
// every single frame. For MP4/MOV we instead read the sample table with
// mp4box.js, then decode the needed span once, in order, with VideoDecoder.

const MP4BOX_SRC = 'vendor/mp4box-0.5.4.min.js';
const DEMUX_CHUNK_BYTES = 4 * 1024 * 1024;
const SAMPLE_READ_WINDOW = 8 * 1024 * 1024;

let mp4boxLoading = null;

/** Load mp4box.js on first use only. */
function loadMp4Box() {
    if (typeof MP4Box !== 'undefined') return Promise.resolve();
    if (!mp4boxLoading) {
        mp4boxLoading = new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = MP4BOX_SRC;
            s.onload = resolve;
            s.onerror = () => { mp4boxLoading = null; reject(new Error('mp4box failed to load')); };
            document.head.appendChild(s);
        });
    }
    return mp4boxLoading;
}

function canUseFastExtraction(file) {
    if (!file || typeof VideoDecoder === 'undefined' || typeof EncodedVideoChunk === 'undefined') return false;
    return /^video\/(mp4|quicktime|x-m4v)$/.test(file.type) || /\.(mp4|m4v|mov)$/i.test(file.name);
}

/** Parse the moov box. Skips over mdat: mp4box tells us the next offset it needs. */
async function demuxMp4(file, isCancelled) {
    await loadMp4Box();
    const mp4 = MP4Box.createFile();
    let info = null;
    let error = null;
    mp4.onReady = (i) => { info = i; };
    mp4.onError = (e) => { error = e; };

    let pos = 0;
    while (!info && !error && pos < file.size) {
        if (isCancelled()) return null;
        const buf = await file.slice(pos, pos + DEMUX_CHUNK_BYTES).arrayBuffer();
        buf.fileStart = pos;
        const next = mp4.appendBuffer(buf);
        pos = (typeof next === 'number' && next > pos) ? next : pos + buf.byteLength;
    }
    if (error) throw new Error(`demux error: ${error}`);
    if (!info) throw new Error('no movie header found');

    const track = info.videoTracks[0];
    if (!track) throw new Error('no video track');
    const trak = mp4.getTrackById(track.id);
    if (!trak || !trak.samples || trak.samples.length === 0) throw new Error('empty sample table');
    return { info, track, trak };
}

/**
 * tkhd matrix → clockwise display rotation in degrees (0/90/180/270),
 * or null for anything that isn't a pure quarter turn.
 */
function matrixRotation(matrix) {
    if (!matrix) return 0;
    const [a, b, , c, d] = Array.from(matrix).map(v => v / 65536);
    const is = (v, t) => Math.abs(v - t) < 1e-3;
    if (is(a, 1) && is(b, 0) && is(c, 0) && is(d, 1)) return 0;
    if (is(a, 0) && is(b, 1) && is(c, -1) && is(d, 0)) return 90;
    if (is(a, -1) && is(b, 0) && is(c, 0) && is(d, -1)) return 180;
    if (is(a, 0) && is(b, -1) && is(c, 1) && is(d, 0)) return 270;
    return null;
}

/** avcC / hvcC / vpcC / av1C payload for VideoDecoder.configure(). */
function codecDescription(trak) {
    for (const entry of trak.mdia.minf.stbl.stsd.entries) {
        const box = entry.avcC || entry.hvcC || entry.vpcC || entry.av1C;
        if (box) {
            const stream = new DataStream(undefined, 0, DataStream.BIG_ENDIAN);
            box.write(stream);
            return new Uint8Array(stream.buffer, 8);   // strip the box header
        }
    }
    return undefined;
}

/**
 * Edit list → offset (seconds) between sample composition time and the
 * presentation time the <video> element uses (e.g. B-frame priming delay).
 */
function editListOffset(trak, movieTimescale) {
    const entries = trak.edts && trak.edts.elst && trak.edts.elst.entries;
    if (!entries || entries.length === 0) return 0;
    const mediaTimescale = trak.mdia.mdhd.timescale;
    let offset = 0;
    for (const e of entries) {
        if (e.media_time === -1) {             // empty edit: delays the track
            offset += e.segment_duration / movieTimescale;
            continue;
        }
        offset -= e.media_time / mediaTimescale;
        break;
    }
    return offset;
}

/** Reads sample payloads through a sliding window instead of one read per sample. */
function createSampleReader(file) {
    let winStart = 0;
    let winBuf = null;
    return async (offset, size) => {
        if (!winBuf || offset < winStart || offset + size > winStart + winBuf.byteLength) {
            winStart = offset;
            winBuf = await file.slice(offset, offset + Math.max(SAMPLE_READ_WINDOW, size)).arrayBuffer();
        }
        return new Uint8Array(winBuf, offset - winStart, size);
    };
}

/**
 * Decode every frame between the keyframe before the first target and the
 * keyframe after the last one, keeping the frame shown at each target time.
 * Throws if the file/codec isn't suitable (caller falls back to seeking);
 * returns null if cancelled.
 */
async function extractFramesWebCodecs(file, { targets, tCtx, temp, isCancelled, onProgress }) {
    const demux = await demuxMp4(file, isCancelled);
    if (!demux) return null;
    const { info, track, trak } = demux;

    // Phone footage carries a rotation matrix the <video> element applies;
    // replicate quarter turns, leave anything else to the seek path.
    const rotation = matrixRotation(track.matrix);
    if (rotation === null) throw new Error('unsupported transform matrix');
    const swap = rotation === 90 || rotation === 270;
    const dispW = swap ? track.video.height : track.video.width;
    const dispH = swap ? track.video.width : track.video.height;
    if (Math.abs(dispW - state.video.width) > 2 || Math.abs(dispH - state.video.height) > 2) {
        throw new Error('display size differs from coded size');
    }

    const config = {
        codec: track.codec,
        codedWidth: track.video.width,
        codedHeight: track.video.height,
        description: codecDescription(trak),
        optimizeForLatency: false,
    };
    const support = await VideoDecoder.isConfigSupported(config).catch(() => null);
    if (!support || !support.supported) throw new Error(`codec ${track.codec} not supported`);

    const timescale = trak.mdia.mdhd.timescale;
    const offset = editListOffset(trak, info.timescale);
    const samples = trak.samples;   // decode order
    const ptsOf = (smp) => smp.cts / timescale + offset;
    const firstT = targets[0];
    const lastT = targets[targets.length - 1];

    // Start at the last keyframe at or before the first target...
    let startIdx = 0;
    for (let i = 0; i < samples.length; i++) {
        if (samples[i].is_sync && ptsOf(samples[i]) <= firstT + 1e-6) startIdx = i;
    }
    // ...and stop at the first keyframe after the last target (all frames before it decode).
    let endIdx = samples.length;
    for (let i = startIdx + 1; i < samples.length; i++) {
        if (samples[i].is_sync && ptsOf(samples[i]) > lastT) { endIdx = i; break; }
    }

    const frames = new Array(targets.length).fill(null);
    let ti = 0;                 // next target to fill
    let held = null;            // latest decoded frame (presentation order)
    let heldBitmap = null;      // its bitmap, once a target needed it
    let decodeError = null;
    let outputChain = Promise.resolve();
    let pendingOutputs = 0;

    const W = temp.width, H = temp.height;
    const bitmapFor = async (frame) => {
        tCtx.save();
        if (rotation === 90) { tCtx.translate(W, 0); tCtx.rotate(Math.PI / 2); }
        else if (rotation === 180) { tCtx.translate(W, H); tCtx.rotate(Math.PI); }
        else if (rotation === 270) { tCtx.translate(0, H); tCtx.rotate(-Math.PI / 2); }
        if (swap) tCtx.drawImage(frame, 0, 0, H, W);
        else tCtx.drawImage(frame, 0, 0, W, H);
        tCtx.restore();
        return createImageBitmap(temp);
    };
    // Every target before time t is shown by the held frame.
    const fillUntil = async (t) => {
        while (ti < targets.length && targets[ti] < t - 1e-6) {
            if (!heldBitmap) heldBitmap = await bitmapFor(held);
            frames[ti++] = heldBitmap;
        }
        onProgress(ti);
    };
    const handleFrame = async (frame) => {
        const t = frame.timestamp / 1e6;
        if (!held) {
            held = frame;
            await fillUntil(t);   // targets before the first frame get the first frame
            return;
        }
        await fillUntil(t);
        held.close();
        held = frame;
        heldBitmap = null;
    };

    const decoder = new VideoDecoder({
        output: (frame) => {
            pendingOutputs++;
            outputChain = outputChain
                .then(() => (isCancelled() || decodeError) ? frame.close() : handleFrame(frame))
                .catch((err) => { decodeError = decodeError || err; frame.close(); })
                .finally(() => { pendingOutputs--; });
        },
        error: (err) => { decodeError = decodeError || err; },
    });

    const cleanup = () => {
        if (decoder.state !== 'closed') decoder.close();
        if (held) held.close();
        held = null;
        closeBitmaps(frames.filter(Boolean));
    };

    try {
        decoder.configure(config);
        const read = createSampleReader(file);
        for (let i = startIdx; i < endIdx; i++) {
            if (isCancelled()) { await outputChain; cleanup(); return null; }
            if (decodeError) throw decodeError;
            // Backpressure: decoder input queue and our bitmap conversion queue
            while (decoder.decodeQueueSize > 8 || pendingOutputs > 4) {
                await new Promise(r => setTimeout(r, 2));
            }
            const smp = samples[i];
            decoder.decode(new EncodedVideoChunk({
                type: smp.is_sync ? 'key' : 'delta',
                timestamp: Math.round(ptsOf(smp) * 1e6),
                duration: Math.round(smp.duration / timescale * 1e6),
                data: await read(smp.offset, smp.size),
            }));
        }
        await decoder.flush();
        await outputChain;
        if (decodeError) throw decodeError;
        if (isCancelled()) { cleanup(); return null; }
        if (!held) throw new Error('decoder produced no frames');
        await fillUntil(Infinity);   // targets after the last frame keep the last frame
        held.close();
        held = null;
        decoder.close();
        return frames;
    } catch (err) {
        await outputChain.catch(() => {});
        cleanup();
        throw err;
    }
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
    updateCacheInfo();
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
    updateExportQualityWarning();
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
            drawTileFrame(ctx, img, t, t.x, t.y, t.w, t.h);
        }
    }

    // Pin overlays — SKIPPED during export (bug fix: pins used to appear in PNG).
    // They follow "Show grid", so a fully frozen grid can be previewed clean.
    if (!state.isExporting) {
        for (const t of state.tiles) {
            if (t.isPinned && el.checkGrid.checked) {
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

/** Draw a tile's source crop. Tile src coords are in source pixels; the cache may be downscaled. */
function drawTileFrame(c, img, tile, dx, dy, dw, dh) {
    const sx = state.frameScaleX, sy = state.frameScaleY;
    c.drawImage(img, tile.srcX * sx, tile.srcY * sy, tile.srcW * sx, tile.srcH * sy, dx, dy, dw, dh);
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
    const t = tileAtEvent(e);
    // Frozen (pinned) tiles can be re-timed in any mode: drag moves their frozen frame.
    // Live tiles scrub only in Standard mode (even while playing).
    if (t && (t.isPinned || anim.mode === 'standard')) {
        state.activeTile = t;
        state.hoverTile = t;
        state.isDragging = true;
        state.scrubbingTile = t;
        state.dragStartX = e.clientX;
        state.dragStartFrame = t.isPinned ? t.frameIndex : t.frameOffset;
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
        if (state.activeTile && state.activeTile.isPinned) {
            const idx = Math.max(0, Math.min(N - 1, newOffset));
            if (state.activeTile.frameIndex !== idx) {
                state.activeTile.frameIndex = idx;
                renderAll();
                updateTimeline();
            }
        } else if (state.activeTile) {
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
        commitEdit();
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
        commitEdit();
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
        commitEdit();
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

/**
 * Time freeze: pin every tile on its own moment, stepping through the range in
 * the given spatial order — a moving subject is frozen across the grid like a
 * chronophotograph. span (0–1) is the fraction of the range covered.
 */
function applyTimeFreeze(order, span) {
    const cols = parseInt(el.inputCols.value, 10);
    const rows = Math.ceil(state.tiles.length / cols);
    const n = state.tiles.length;
    const N = state.totalFrames;
    if (n === 0 || N === 0) return;

    let ranks;
    if (order === 'shuffle') {
        ranks = Array.from({ length: n }, (_, i) => i);
        for (let i = n - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [ranks[i], ranks[j]] = [ranks[j], ranks[i]];
        }
    } else {
        ranks = state.tiles.map((_, i) => {
            const c = i % cols, r = Math.floor(i / cols);
            if (order === 'rl') return n - 1 - i;
            if (order === 'tb') return c * rows + r;
            return i;   // 'lr'
        });
    }

    const step = span * (N - 1) / Math.max(1, n - 1);
    state.tiles.forEach((t, i) => {
        t.isPinned = true;
        t.frameIndex = Math.max(0, Math.min(N - 1, Math.round(ranks[i] * step)));
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
    setCancellable(null);
}

// Long tasks (extraction, export) can register a cancel handler for the overlay button / Esc
let cancelHandler = null;

function setCancellable(fn) {
    cancelHandler = fn;
    el.loadingOverlay.classList.toggle('cancellable', !!fn);
}

function cancelCurrentTask() {
    const fn = cancelHandler;
    setCancellable(null);
    if (fn) fn();
}

el.btnCancelTask.addEventListener('click', cancelCurrentTask);

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
    el.btnProjectSave.disabled = false;
    el.inputFreezeSpan.disabled = false;
    el.btnUnfreeze.disabled = false;
    document.querySelectorAll('.freeze-btn').forEach(b => { b.disabled = false; });
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
 * Export canvas size for a resolution tier, keeping the preview canvas aspect
 * (hardcoded 16:9 dims would stretch portrait or square content).
 * aspect defaults to the current canvas; even = H.264-safe dimensions.
 */
function exportSize(res, { even = true, aspect = canvas.width / canvas.height } = {}) {
    let w, h;
    if (res === 'preview') {
        w = Math.round(aspect >= 1 ? MAX_CANVAS_DIM : MAX_CANVAS_DIM * aspect);
        h = Math.round(aspect >= 1 ? MAX_CANVAS_DIM / aspect : MAX_CANVAS_DIM);
        if (state.tiles.length) { w = canvas.width; h = canvas.height; }
    } else {
        const longSide = { '720': 1280, '1080': 1920, '4k': 3840 }[res] || 1920;
        if (aspect >= 1) { w = longSide; h = Math.round(longSide / aspect); }
        else             { h = longSide; w = Math.round(longSide * aspect); }
    }
    return even ? ensureEvenDimensions(w, h) : { w, h };
}

/**
 * Ensure canvas dimensions are even (required for H.264 compliance).
 */
function ensureEvenDimensions(w, h) {
    return {
        w: w + (w % 2),
        h: h + (h % 2)
    };
}

/**
 * Export video — always stops preview first.
 */
function exportFilename(ext) {
    const res = el.selectRes.value;
    const fps = parseInt(el.selectFps.value, 10);
    const dur = parseFloat(el.inputDuration.value);
    const resLabel = res === 'preview' ? `${canvas.width}x${canvas.height}` : res;
    return `PanoTile-${resLabel}-${fps}fps-${dur}s-${Date.now()}.${ext}`;
}

function cancelledError() {
    return new DOMException('Export cancelled', 'AbortError');
}

async function exportVideo() {
    if (state.isExporting || !state.ready || extractionRunning) return;

    const path = getExportPath();

    // Stream long exports straight to disk where supported (Chromium File System
    // Access API) so the MP4 never has to fit in memory. The picker must open
    // first, while the click still counts as a user gesture.
    let fileHandle = null;
    if (path === 'A' && typeof window.showSaveFilePicker === 'function') {
        try {
            fileHandle = await window.showSaveFilePicker({
                suggestedName: exportFilename('mp4'),
                types: [{ description: 'MP4 video', accept: { 'video/mp4': ['.mp4'] } }],
            });
        } catch (err) {
            if (err.name === 'AbortError') return;   // user closed the save dialog
            fileHandle = null;                        // picker unavailable — in-memory download
        }
    }

    // Stop preview before export
    pause();

    state.isExporting = true;
    updateStatus();
    // Path B records in real time and stalls in background tabs
    showLoading(path === 'A' ? 'EXPORTING...' : 'EXPORTING IN REAL TIME — KEEP THIS TAB VISIBLE');
    const abort = { cancelled: false };
    setCancellable(() => {
        abort.cancelled = true;
        if (el.loadingText) el.loadingText.innerText = 'CANCELLING...';
    });

    console.log(`Export path: ${path} (${path === 'A' ? 'WebCodecs+mp4-muxer' : 'MediaRecorder fallback'}${fileHandle ? ', streaming to disk' : ''})`);

    try {
        if (path === 'A') {
            await exportVideoWebCodecs(abort, fileHandle);
        } else {
            await exportVideoMediaRecorder(abort);
        }
    } catch (err) {
        if (err && err.name === 'AbortError') {
            showToast('Export cancelled', 'info');
        } else {
            console.error('Export failed:', err);
            showToast(`Export failed: ${err.message || 'Unknown error'}`);
        }
    } finally {
        hideLoading();
        state.isExporting = false;
        updateStatus();
    }
}

/**
 * Path A: WebCodecs + mp4-muxer (Primary - MP4 H.264)
 */
async function exportVideoWebCodecs(abort, fileHandle = null) {
    const fps = parseInt(el.selectFps.value, 10);
    const dur = parseFloat(el.inputDuration.value);
    const res = el.selectRes.value;
    const totalFrames = Math.ceil(dur * fps);

    // Determine export canvas size
    const { w: expW, h: expH } = exportSize(res);

    const bitrate = exportBitrate(expW, expH, fps);

    // Create export canvas
    const expCanvas = document.createElement('canvas');
    expCanvas.width = expW;
    expCanvas.height = expH;
    const expCtx = expCanvas.getContext('2d');
    expCtx.imageSmoothingQuality = 'high';
    // Fresh phase state so the export is deterministic and never mutates the preview
    const exportPhases = createInitialPhases();

    // Setup mp4-muxer — disk stream (moov at the end) or in-memory buffer (moov first)
    const writable = fileHandle ? await fileHandle.createWritable() : null;
    const muxer = new Mp4Muxer.Muxer({
        target: writable
            ? new Mp4Muxer.FileSystemWritableFileStreamTarget(writable)
            : new Mp4Muxer.ArrayBufferTarget(),
        video: {
            codec: 'avc',
            width: expW,
            height: expH,
            bitrate: bitrate
        },
        fastStart: writable ? false : 'in-memory'
    });

    // Capture encoder errors via variable — throwing inside WebCodecs callbacks
    // does NOT propagate to the outer async try/catch; it only closes the encoder.
    let encoderError = null;
    let frameCount = 0;
    const encoder = new VideoEncoder({
        output: (chunk, meta) => muxer.addVideoChunk(chunk, meta),
        error: (err) => { encoderError = err; }
    });

    try {
        // Fail fast (before rendering anything) if the GPU/browser can't encode this size
        const encoderConfig = await chooseEncoderConfig(expW, expH, fps, bitrate);
        if (!encoderConfig) {
            throw new Error(`H.264 ${expW}×${expH} @ ${fps}fps is not supported by this browser — try a lower resolution`);
        }
        console.info(`Encoder: ${encoderConfig.codec}, ${(bitrate / 1e6).toFixed(1)} Mbps`);
        encoder.configure(encoderConfig);

        // Render each frame
        for (let i = 0; i < totalFrames; i++) {
            if (abort.cancelled) throw cancelledError();

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
        if (abort.cancelled) throw cancelledError();
        muxer.finalize();

        if (writable) {
            await writable.close();   // waits for every queued write to land on disk
            const file = await fileHandle.getFile().catch(() => null);
            showToast(`Saved ${fileHandle.name}${file ? ` (${formatBytes(file.size)})` : ''}`, 'info');
            return;
        }

        // Get the buffer
        const buffer = muxer.target.buffer;
        const blob = new Blob([buffer], { type: 'video/mp4' });
        const filename = exportFilename('mp4');

        // Download
        downloadBlob(blob, filename);

        showToast(`Exported ${filename} (${formatBytes(blob.size)})`, 'info');
    } catch (err) {
        if (encoder.state !== 'closed') encoder.close();
        // Discard the partially written file contents
        if (writable) await writable.abort().catch(() => {});
        throw err;
    }
}

/**
 * Bitrate from pixel rate (0.12 bits per pixel per frame), never below the old
 * per-tier defaults and capped at 100 Mbps. 4K60 ≈ 55 Mbps, 1080p30 = 10 Mbps.
 */
function exportBitrate(w, h, fps) {
    const floor = w * h > 2.1e6 ? 20e6 : w * h > 0.93e6 ? 10e6 : 6e6;
    return Math.round(Math.min(100e6, Math.max(floor, w * h * fps * 0.12)));
}

// H.264 levels: max macroblocks per frame / per second (ITU-T H.264 Table A-1)
const AVC_LEVELS = [
    { hex: '1f', fs: 3600, mbps: 108000 },     // 3.1
    { hex: '28', fs: 8192, mbps: 245760 },     // 4.0
    { hex: '2a', fs: 8704, mbps: 522240 },     // 4.2
    { hex: '33', fs: 36864, mbps: 983040 },    // 5.1
    { hex: '34', fs: 36864, mbps: 2073600 },   // 5.2
    { hex: '3c', fs: 139264, mbps: 4177920 },  // 6.0
];

/**
 * Best supported H.264 config: High profile first (much better quality per bit
 * than Baseline), then Main, then Baseline; level picked from frame size AND
 * frame rate (4K60 needs 5.2, not 5.1). Returns null if nothing is supported.
 */
async function chooseEncoderConfig(w, h, fps, bitrate) {
    const fs = Math.ceil(w / 16) * Math.ceil(h / 16);
    const level = AVC_LEVELS.find(l => fs <= l.fs && fs * fps <= l.mbps) || AVC_LEVELS[AVC_LEVELS.length - 1];
    for (const profile of ['6400', '4d00', '4200']) {   // High, Main, Baseline
        const config = { codec: `avc1.${profile}${level.hex}`, width: w, height: h, bitrate, framerate: fps };
        const support = await VideoEncoder.isConfigSupported(config).catch(() => null);
        if (support && support.supported) return config;
    }
    return null;
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
            drawTileFrame(expCtx, img, tile, tile.x * scaleX, tile.y * scaleY, tile.w * scaleX, tile.h * scaleY);
        }
    }
}

/**
 * Path B: MediaRecorder (Fallback - WebM)
 */
async function exportVideoMediaRecorder(abort) {
    const fps = parseInt(el.selectFps.value, 10);
    const dur = parseFloat(el.inputDuration.value);
    const res = el.selectRes.value;
    const totalFrames = Math.ceil(dur * fps);

    // Determine export canvas size
    const { w: expW, h: expH } = exportSize(res);

    // Create export canvas
    const expCanvas = document.createElement('canvas');
    expCanvas.width = expW;
    expCanvas.height = expH;
    const expCtx = expCanvas.getContext('2d');
    expCtx.imageSmoothingQuality = 'high';
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

    const recorder = new MediaRecorder(stream, {
        ...(mimeType ? { mimeType } : {}),
        videoBitsPerSecond: exportBitrate(expW, expH, fps),
    });
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
    const stopRecording = async () => {
        const stopPromise = new Promise(r => { recorder.onstop = r; });
        if (recorder.state !== 'inactive') recorder.stop();
        else return;
        await stopPromise;
    };
    for (let i = 0; i < totalFrames; i++) {
        if (abort.cancelled) {
            await stopRecording();
            for (const track of stream.getTracks()) track.stop();
            throw cancelledError();
        }
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

    // onstop is set before stop() to avoid a race
    await stopRecording();
    for (const track of stream.getTracks()) track.stop();

    const blob = new Blob(chunks, { type: mimeType || 'video/webm' });
    const filename = exportFilename('webm');

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
    const { w: expW, h: expH } = exportSize(res, { even: false });

    const expCanvas = document.createElement('canvas');
    expCanvas.width = expW;
    expCanvas.height = expH;
    const expCtx = expCanvas.getContext('2d');
    expCtx.imageSmoothingQuality = 'high';

    const scaleX = expW / canvas.width;
    const scaleY = expH / canvas.height;

    for (const tile of state.tiles) {
        const img = state.frames[tile.frameIndex];
        if (img) {
            drawTileFrame(expCtx, img, tile, tile.x * scaleX, tile.y * scaleY, tile.w * scaleX, tile.h * scaleY);
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
    commitEdit();
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
    commitEdit();
});

el.inputSpatialAmt.addEventListener('input', () => {
    el.spatialAmtValue.innerText = el.inputSpatialAmt.value;
    if (el.checkSpatialShuffle.checked) {
        resetSpatialShuffle();
        applySpatialShuffle(parseInt(el.inputSpatialAmt.value));
        if (!anim.playing) renderAll();
    }
});
el.inputSpatialAmt.addEventListener('change', commitEdit);

// Pattern
el.selectPattern.addEventListener('change', () => {
    el.patternAmtRow.style.display = el.selectPattern.value === 'random' ? 'flex' : 'none';
});

el.btnPatternApply.addEventListener('click', () => {
    const pattern = el.selectPattern.value;
    if (pattern === 'none') return;
    const amount = parseInt(el.inputPatternAmt.value);
    applyBlockPattern(pattern, amount);
    commitEdit();
});

el.btnPatternClear.addEventListener('click', () => {
    clearAllPins();
    commitEdit();
});

// Time freeze
el.inputFreezeSpan.addEventListener('input', () => {
    el.freezeSpanValue.innerText = el.inputFreezeSpan.value;
});
document.querySelectorAll('.freeze-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        if (!state.ready) return;
        applyTimeFreeze(btn.dataset.freeze, parseInt(el.inputFreezeSpan.value, 10) / 100);
        commitEdit();
    });
});
el.btnUnfreeze.addEventListener('click', () => {
    clearAllPins();
    commitEdit();
});

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
    updateCacheInfo();
    if (state.ready) reextractFrames();
});

// Auto frame quality follows the export resolution — re-extract if the target size moved
el.selectRes.addEventListener('change', () => {
    updateCacheInfo();
    if (state.ready && el.selectCache.value === 'auto') {
        const target = computeCacheSize(state.totalFrames).w;
        const current = Math.round(state.video.width * state.frameScaleX);
        if (Math.abs(target - current) / current > 0.05) {
            reextractFrames();
            return;
        }
    }
    updateExportQualityWarning();
});

// Frame quality defines the cache resolution — re-extract on change
el.selectCache.addEventListener('change', () => {
    updateCacheInfo();
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
    commitEdit();
});

// Loop mode dropdown
if (el.selectLoop) {
    el.selectLoop.addEventListener('change', () => {
        anim.loopMode = el.selectLoop.value;
        // Sync checkbox
        el.checkLoop.checked = anim.loopMode === 'wrap';
        commitEdit();
    });
}

// Stutter slider
if (el.inputStutter) {
    el.inputStutter.addEventListener('input', () => {
        anim.stutter = parseInt(el.inputStutter.value, 10);
        if (el.stutterValue) el.stutterValue.innerText = anim.stutter;
    });
    el.inputStutter.addEventListener('change', commitEdit);
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
    commitEdit();
}

el.checkSquare.addEventListener('change', rebuildGrid);
el.inputCols.addEventListener('change', rebuildGrid);
el.inputRows.addEventListener('change', () => {
    if (!el.checkSquare.checked) rebuildGrid();
});

/* ==========================================================
   PROJECT — save / open, undo / redo, autosave
   ========================================================== */

const PROJECT_VERSION = 1;
const AUTOSAVE_KEY = 'panotile_autosave_v1';
const HISTORY_LIMIT = 100;
const MAX_PROJECT_FILE_BYTES = 20 * 1024 * 1024;

const editHistory = { undo: [], redo: [], current: null };
let pendingProject = null;   // project opened before its video was loaded
let autosaveTimer = null;

/** Everything needed to rebuild the edit (not the video itself). */
function serializeProject() {
    return {
        app: 'PanoTile',
        version: PROJECT_VERSION,
        video: {
            name: state.video.name,
            size: state.video.size,
            duration: state.video.duration,
            width: state.video.width,
            height: state.video.height,
        },
        range: { in: state.range.in, out: state.range.out },
        totalFrames: state.totalFrames,
        settings: {
            cols: parseInt(el.inputCols.value, 10),
            rows: parseInt(el.inputRows.value, 10),
            square: el.checkSquare.checked,
            fps: el.selectFps.value,
            quality: el.selectCache.value,
            mode: anim.mode,
            loopMode: anim.loopMode,
            stutter: anim.stutter,
            spatialShuffle: el.checkSpatialShuffle.checked,
            spatialAmt: parseInt(el.inputSpatialAmt.value, 10),
            showGrid: el.checkGrid.checked,
            exportRes: el.selectRes.value,
            exportDuration: parseFloat(el.inputDuration.value),
        },
        tileOffsets: anim.tileOffsets.slice(),
        tiles: state.tiles.map(t => ({
            srcX: t.srcX,
            srcY: t.srcY,
            pinned: t.isPinned,
            frame: t.frameIndex,
            offset: t.frameOffset,
        })),
    };
}

const clampNum = (v, min, max, def) => {
    const n = Number(v);
    return Number.isFinite(n) ? Math.max(min, Math.min(max, n)) : def;
};
const pickOption = (sel, v, def) =>
    [...sel.options].some(o => o.value === String(v)) ? String(v) : def;

/**
 * Validate and normalise a project object. Project files are untrusted input:
 * every field is type-checked and clamped; unknown fields are dropped.
 */
function parseProject(raw) {
    if (!raw || typeof raw !== 'object' || raw.app !== 'PanoTile') {
        throw new Error('not a PanoTile project file');
    }
    if (raw.version !== PROJECT_VERSION) {
        throw new Error(`unsupported project version (${raw.version})`);
    }
    const st = raw.settings && typeof raw.settings === 'object' ? raw.settings : {};
    const v = raw.video && typeof raw.video === 'object' ? raw.video : {};
    const tilesIn = Array.isArray(raw.tiles) ? raw.tiles.slice(0, 10000) : [];
    const maxF = MAX_SOURCE_FRAMES;
    const rIn = clampNum(raw.range?.in, 0, 1e7, 0);

    return {
        video: {
            name: typeof v.name === 'string' ? v.name.slice(0, 500) : '',
            size: clampNum(v.size, 0, Number.MAX_SAFE_INTEGER, 0),
            duration: clampNum(v.duration, 0, 1e7, 0),
            width: Math.round(clampNum(v.width, 0, 1e5, 0)),
            height: Math.round(clampNum(v.height, 0, 1e5, 0)),
        },
        range: { in: rIn, out: clampNum(raw.range?.out, rIn, 1e7, rIn) },
        totalFrames: Math.round(clampNum(raw.totalFrames, 1, maxF, 1)),
        settings: {
            cols: Math.round(clampNum(st.cols, 1, 30, 4)),
            rows: Math.round(clampNum(st.rows, 1, 100, 4)),
            square: st.square === true,
            fps: pickOption(el.selectFps, st.fps, el.selectFps.value),
            quality: pickOption(el.selectCache, st.quality, 'auto'),
            mode: pickOption(el.selectMode, st.mode, 'standard'),
            loopMode: pickOption(el.selectLoop, st.loopMode, 'wrap'),
            stutter: Math.round(clampNum(st.stutter, 1, 6, 1)),
            spatialShuffle: st.spatialShuffle === true,
            spatialAmt: Math.round(clampNum(st.spatialAmt, 0, 10, 5)),
            showGrid: st.showGrid !== false,
            exportRes: pickOption(el.selectRes, st.exportRes, el.selectRes.value),
            exportDuration: clampNum(st.exportDuration, 1, MAX_EXPORT_DURATION, 5),
        },
        tileOffsets: (Array.isArray(raw.tileOffsets) ? raw.tileOffsets.slice(0, 10000) : [])
            .map(o => clampNum(o, 0, 1, 0)),
        tiles: tilesIn.map(t => ({
            srcX: clampNum(t?.srcX, 0, 1e5, 0),
            srcY: clampNum(t?.srcY, 0, 1e5, 0),
            pinned: t?.pinned === true,
            frame: Math.round(clampNum(t?.frame, 0, maxF - 1, 0)),
            offset: Math.round(clampNum(t?.offset, -(maxF - 1), maxF - 1, 0)),
        })),
    };
}

/**
 * Apply grid, animation and tile state from a parsed project to the loaded
 * video. Range/fps/quality are NOT touched here (they need a re-extraction —
 * see applyProjectFile); frame indices are remapped onto the current frames.
 */
function applyProjectState(p) {
    const wasPlaying = anim.playing;
    const st = p.settings;

    el.inputCols.value = st.cols;
    el.inputRows.value = st.rows;
    el.checkSquare.checked = st.square;
    anim.mode = st.mode;
    el.selectMode.value = st.mode;
    anim.loopMode = st.loopMode;
    el.selectLoop.value = st.loopMode;
    el.checkLoop.checked = st.loopMode === 'wrap';
    anim.stutter = st.stutter;
    el.inputStutter.value = st.stutter;
    el.stutterValue.innerText = st.stutter;
    el.checkSpatialShuffle.checked = st.spatialShuffle;
    el.inputSpatialAmt.value = st.spatialAmt;
    el.spatialAmtValue.innerText = st.spatialAmt;
    el.checkGrid.checked = st.showGrid;
    el.selectRes.value = st.exportRes;
    el.inputDuration.value = st.exportDuration;
    el.durationValue.innerText = st.exportDuration.toFixed(1);

    initProject();   // rebuilds tiles for this grid; setMode() may auto-play
    pause();

    if (p.tiles.length === state.tiles.length) {
        // Source crops are in source pixels — only meaningful for a same-size video
        const sameFrameSize = p.video.width === state.video.width && p.video.height === state.video.height;
        state.tiles.forEach((t, i) => {
            const s = p.tiles[i];
            if (sameFrameSize) { t.srcX = s.srcX; t.srcY = s.srcY; }
            t.isPinned = s.pinned;
            t.frameIndex = s.frame;
            t.frameOffset = s.offset;
        });
        remapTileFrames(p.range, p.totalFrames);
    }
    if (p.tileOffsets.length === state.tiles.length) {
        anim.tileOffsets = p.tileOffsets.slice();
        anim.tilePhases = createInitialPhases();
    }
    anim.elapsed = 0;
    if (wasPlaying && anim.mode !== 'standard') play();
    rememberGrid();
    renderAll();
    updateStatus();
    updateTimeline();
}

/** Apply a project file, re-extracting frames first if range/fps/quality differ. */
async function applyProjectFile(p) {
    const st = p.settings;
    const dur = state.video.duration;
    const rIn = Math.max(0, Math.min(dur - MIN_RANGE_SEC, p.range.in));
    const rOut = Math.max(rIn + MIN_RANGE_SEC, Math.min(dur, p.range.out));
    const needsExtract = st.fps !== el.selectFps.value
        || st.quality !== el.selectCache.value
        || Math.abs(rIn - state.range.in) > 1e-6
        || Math.abs(rOut - state.range.out) > 1e-6;

    if (needsExtract) {
        el.selectFps.value = st.fps;
        el.selectCache.value = st.quality;
        state.range = { in: rIn, out: rOut };
        renderRangeUI();
        pause();
        state.ready = false;
        const ok = await extractFrames(state.video.probe, {
            keepProject: true,
            prevRange: lastExtracted && lastExtracted.range,
        });
        if (!ok) return false;
    }
    applyProjectState(p);
    resetHistory();
    scheduleAutosave();
    return true;
}

async function offerApplyProject(p) {
    const sameVideo = p.video.name === state.video.name && p.video.size === state.video.size;
    if (!sameVideo) {
        const ok = await confirm(
            'Different video',
            `This project was made for "${p.video.name || 'unknown'}". Apply it to the current video anyway?`,
            'Apply'
        );
        if (!ok) return;
    }
    if (await applyProjectFile(p)) showToast('Project loaded', 'info');
}

/* --- Undo / redo: snapshots of the serialised project --- */

function resetHistory() {
    editHistory.undo = [];
    editHistory.redo = [];
    editHistory.current = state.ready ? JSON.stringify(serializeProject()) : null;
    updateHistoryUI();
}

/** Call after every user edit; records an undo step only if something changed. */
function commitEdit() {
    if (!state.ready) return;
    const snap = JSON.stringify(serializeProject());
    if (snap === editHistory.current) return;
    if (editHistory.current) {
        editHistory.undo.push(editHistory.current);
        if (editHistory.undo.length > HISTORY_LIMIT) editHistory.undo.shift();
    }
    editHistory.current = snap;
    editHistory.redo = [];
    updateHistoryUI();
    scheduleAutosave();
}

function canStepHistory() {
    return state.ready && !state.isExporting && !extractionRunning && !state.isDragging;
}

function undo() {
    if (!editHistory.undo.length || !canStepHistory()) return;
    editHistory.redo.push(editHistory.current);
    editHistory.current = editHistory.undo.pop();
    applyProjectState(parseProject(JSON.parse(editHistory.current)));
    updateHistoryUI();
    scheduleAutosave();
}

function redo() {
    if (!editHistory.redo.length || !canStepHistory()) return;
    editHistory.undo.push(editHistory.current);
    editHistory.current = editHistory.redo.pop();
    applyProjectState(parseProject(JSON.parse(editHistory.current)));
    updateHistoryUI();
    scheduleAutosave();
}

function updateHistoryUI() {
    el.btnUndo.disabled = editHistory.undo.length === 0;
    el.btnRedo.disabled = editHistory.redo.length === 0;
}

/* --- Autosave: one slot in localStorage, keyed to the video --- */

function videoKey() {
    return `${state.video.name}|${state.video.size}|${state.video.duration.toFixed(3)}`;
}

function scheduleAutosave() {
    clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(() => {
        if (!state.ready) return;
        try {
            localStorage.setItem(AUTOSAVE_KEY, JSON.stringify({
                key: videoKey(),
                savedAt: Date.now(),
                project: serializeProject(),
            }));
        } catch (_) { /* storage full or blocked — autosave is best effort */ }
    }, 800);
}

function readAutosave() {
    try {
        const raw = JSON.parse(localStorage.getItem(AUTOSAVE_KEY));
        if (raw && raw.key === videoKey()) return raw;
    } catch (_) {}
    return null;
}

/** Called by extractFrames when frames are in place. */
async function onFramesReady({ fresh }) {
    if (!fresh) {
        // Re-extraction (range/fps/quality): not undoable, but keep earlier steps usable
        editHistory.current = JSON.stringify(serializeProject());
        scheduleAutosave();
        return;
    }
    resetHistory();

    if (pendingProject) {
        const p = pendingProject;
        pendingProject = null;
        await offerApplyProject(p);
        return;
    }

    const saved = readAutosave();
    if (!saved) return;
    let p;
    try { p = parseProject(saved.project); } catch (_) { return; }
    const when = new Date(saved.savedAt).toLocaleString();
    const ok = await confirm('Restore session?', `Found autosaved edits for this video from ${when}.`, 'Restore');
    if (ok && await applyProjectFile(p)) showToast('Session restored', 'info');
}

/* --- Save / open buttons --- */

el.btnProjectSave.addEventListener('click', () => {
    if (!state.ready) return;
    const data = { ...serializeProject(), savedAt: new Date().toISOString() };
    const base = state.video.name.replace(/\.[^.]+$/, '').replace(/[^\w.-]+/g, '_').slice(0, 80) || 'project';
    downloadBlob(new Blob([JSON.stringify(data)], { type: 'application/json' }), `${base}.panotile.json`);
    showToast('Project saved', 'info');
});

el.btnProjectOpen.addEventListener('click', () => el.projectUpload.click());

el.projectUpload.addEventListener('change', async (e) => {
    const f = e.target.files?.[0];
    e.target.value = '';
    if (!f) return;
    if (f.size > MAX_PROJECT_FILE_BYTES) {
        showToast('Project file too large');
        return;
    }
    let p;
    try {
        p = parseProject(JSON.parse(await f.text()));
    } catch (err) {
        showToast(`Invalid project: ${err.message}`);
        return;
    }
    if (state.isExporting || extractionRunning) {
        showToast('Wait for the current task to finish');
        return;
    }
    if (!state.ready) {
        pendingProject = p;
        showToast(`Project ready — now load the video "${p.video.name}"`, 'info');
        return;
    }
    await offerApplyProject(p);
});

el.btnUndo.addEventListener('click', undo);
el.btnRedo.addEventListener('click', redo);

/* ==========================================================
   KEYBOARD SHORTCUTS
   ========================================================== */

document.addEventListener('keydown', (e) => {
    // Esc cancels a running extraction/export
    if (e.key === 'Escape' && cancelHandler) {
        cancelCurrentTask();
        return;
    }
    // Ignore if user is typing in an input (native undo there)
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;
    // No shortcuts behind a modal, while loading or while exporting
    if ($('tutorialOverlay')?.classList.contains('visible')) return;
    if (el.confirmDialog.classList.contains('visible')) return;
    if (el.loadingOverlay.classList.contains('visible') || state.isExporting) return;

    // Undo / redo: Cmd/Ctrl+Z, Shift+Cmd/Ctrl+Z, Ctrl+Y
    const mod = e.metaKey || e.ctrlKey;
    const k = e.key.toLowerCase();
    if (mod && !e.altKey && (k === 'z' || k === 'y')) {
        e.preventDefault();
        if (k === 'y' || e.shiftKey) redo(); else undo();
        return;
    }
    // Leave other browser/OS shortcuts alone (Cmd+S, Ctrl+R, …)
    if (mod || e.altKey) return;
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
            commitEdit();
            break;
        case 'k': // Linear R->L
            el.selectMode.value = 'linear-rl';
            setMode('linear-rl');
            commitEdit();
            break;
        case 's': // Shuffle
            el.selectMode.value = 'temporal-shuffle';
            setMode('temporal-shuffle');
            commitEdit();
            break;
        case 'p': // Perlin flow
            el.selectMode.value = 'perlin-flow';
            setMode('perlin-flow');
            commitEdit();
            break;
        case 'd': // Drunk walk
            el.selectMode.value = 'drunk';
            setMode('drunk');
            commitEdit();
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

