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

// Per-frame RAM estimate (empirical: JPEG 0.72 quality ≈ these sizes)
const RAM_PER_FRAME_BY_RES = {
    'preview': 0.12,  // MB per frame, preview canvas
    '720':     0.18,
    '1080':    0.30,
    '4k':      0.95
};

const state = {
    video: { width: 0, height: 0, duration: 0, name: '', size: 0, blobUrl: null, probe: null },
    range: { in: 0, out: 0 },
    totalFrames: 0,
    extractedWith: null,             // { rangeIn, rangeOut, frameCount } — to detect when regenerate is needed
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
    ramEstimate: $('ramEstimate'),
    ramLabel: $('ramLabel'),
    ramFill: $('ramFill'),
    ramValue: $('ramValue'),
    btnGenerate: $('btnGenerate'),
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
    if (!file.type.match(/^video\/(mp4|webm|quicktime)$/)) {
        showToast(`Unsupported format: ${file.type || 'unknown'}`);
        return;
    }

    // Clean previous blob URL to avoid memory leak
    if (state.video.blobUrl) URL.revokeObjectURL(state.video.blobUrl);

    state.video.name = file.name;
    state.video.size = file.size;
    state.video.blobUrl = URL.createObjectURL(file);

    loadVideo(state.video.blobUrl);
}

function loadVideo(src) {
    const probe = document.createElement('video');
    probe.preload = 'metadata';
    probe.muted = true;
    probe.src = src;

    const timeout = setTimeout(() => {
        showToast('Video metadata timeout. Try a different file.');
    }, 8000);

    probe.onloadedmetadata = () => {
        clearTimeout(timeout);

        if (!isFinite(probe.duration) || probe.duration <= 0) {
            showToast('Video has invalid duration metadata.');
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
        clearTimeout(timeout);
        showToast('Could not read video file.');
    };
}

function updateSourceCard() {
    const v = state.video;
    el.sourceCard.classList.remove('empty');
    el.sourceCard.innerHTML = `
        <div class="source-thumb">▶</div>
        <div class="source-meta">
            <div class="source-name">${v.name}</div>
            <div class="source-specs">${v.width}×${v.height} · ${formatBytes(v.size)} · ${v.duration.toFixed(1)}s</div>
        </div>
        <button class="btn-change" id="btnChangeVideo">Change</button>
    `;
    $('btnChangeVideo').onclick = (e) => {
        e.stopPropagation();
        el.videoUpload.click();
    };
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

async function extractFrames(probe) {
    state.frames = [];
    const targetCount = computeSourceFrameCount();
    state.totalFrames = targetCount;
    showLoading(`EXTRACTING ${targetCount} FRAMES`);

    const temp = document.createElement('canvas');
    temp.width = state.video.width;
    temp.height = state.video.height;
    const tCtx = temp.getContext('2d');

    const rangeDur = state.range.out - state.range.in;

    for (let i = 0; i < targetCount; i++) {
        try {
            const t = state.range.in + (i / Math.max(1, targetCount - 1)) * rangeDur;
            await seekAndCapture(probe, tCtx, temp, t);
        } catch (err) {
            console.warn(`Frame ${i} failed:`, err);
            if (state.frames.length > 0) {
                state.frames.push(state.frames[state.frames.length - 1]);
            } else {
                hideLoading();
                showToast('Frame extraction failed. Try a different video.');
                return;
            }
        }
        el.progressFill.style.width = `${((i + 1) / targetCount) * 100}%`;
    }

    hideLoading();
    canvas.style.display = 'block';
    el.emptyState.style.display = 'none';
    el.timeline.classList.add('visible');
    state.ready = true;
    state.extractedWith = {
        rangeIn: state.range.in,
        rangeOut: state.range.out,
        frameCount: state.totalFrames
    };
    enableControls();
    initProject();
    updateStatus();
    updateRamEstimate();
    markRegenerateStatus();
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
                const bitmap = await createImageBitmap(temp);
                state.frames.push(bitmap);
                resolve();
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
    try {
        const THUMB_COUNT = 10;
        const THUMB_W = 80;
        const THUMB_H = 40;
        const vid = document.createElement('video');
        vid.muted = true;
        vid.preload = 'metadata';
        vid.src = state.video.blobUrl;
        await new Promise(r => {
            if (vid.readyState >= 1) r();
            else vid.onloadedmetadata = r;
        });

        const stripCanvas = document.createElement('canvas');
        stripCanvas.width = THUMB_W * THUMB_COUNT;
        stripCanvas.height = THUMB_H;
        const sctx = stripCanvas.getContext('2d');

        for (let i = 0; i < THUMB_COUNT; i++) {
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
        el.rangeStrip.style.backgroundImage = `url(${stripCanvas.toDataURL('image/jpeg', 0.6)})`;
        el.rangeStrip.style.backgroundSize = '100% 100%';
    } catch (e) {
        // Thumbstrip failure is non-fatal — just leave it empty
        console.warn('Thumbstrip failed:', e);
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

[el.rangeHandleL, el.rangeHandleR].forEach((h, idx) => {
    h.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        rangeDragging = idx === 0 ? 'L' : 'R';
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
    updateRamEstimate();
    markRegenerateStatus();
});

window.addEventListener('pointerup', () => {
    if (rangeDragging) {
        rangeDragging = null;
        // Auto-clamp export duration to range if it now exceeds it
        const rd = state.range.out - state.range.in;
        const exportDur = parseFloat(el.inputDuration.value);
        // Allow export longer than range (time-dilation effect) — no auto-clamp.
        // Just update estimator; user triggers re-extraction via Regenerate.
        updateRamEstimate();
    }
});

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
                isPinned: false,
                id: r * cols + c
            });
        }
    }
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
function computeFrameIndex(tile, outputFrame, cols, N, rate, elapsed, tileOffsets, tilePhases, i) {
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
            const rawPhase = stutterFrame;
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
            tilePhases[i] = ((tilePhases[i] + v * rate * (1/60)) % N + N) % N;
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
    ctx.shadowColor = 'rgba(255, 46, 136, 0.5)';
    ctx.shadowBlur = 4;
    ctx.strokeStyle = 'rgba(255, 46, 136, 0.95)';
    ctx.lineWidth = 2;
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
        state.dragStartFrame = t.frameIndex;
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
        let idx = state.dragStartFrame + Math.floor(px / sensitivity);
        idx = Math.max(0, Math.min(state.totalFrames - 1, idx));
        if (state.activeTile && state.activeTile.frameIndex !== idx) {
            state.activeTile.frameIndex = idx;
            renderAll();
            updateTimeline();
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
            anim.tileOffsets, anim.tilePhases, i
        );
    });

    // Auto-pause for standard mode with hold loop when reaching end
    if (anim.mode === 'standard' && anim.loopMode === 'hold') {
        const rawPhase = t * rate;
        if (rawPhase >= N - 1) pause();
    }
}

/* ==========================================================
   SPATIAL SHUFFLE
   ========================================================== */

function applySpatialShuffle(amount) {
    // Spatial Shuffle: swap frameIndex values between tiles.
    // amount (0-10): 0 = no shuffle, 10 = full shuffle.
    // The shuffle intensity determines how many swaps to perform.
    const n = state.tiles.length;
    if (n < 2) return;

    // Collect current frame indices
    const indices = state.tiles.map(t => t.frameIndex);

    // Fisher-Yates shuffle with intensity control
    // amount=10 → full shuffle (n swaps), amount=5 → 50% swaps
    const swapCount = Math.round((amount / 10) * n);

    for (let s = 0; s < swapCount; s++) {
        const a = Math.floor(Math.random() * n);
        const b = Math.floor(Math.random() * n);
        if (a !== b) {
            // Swap frame indices
            const tmp = indices[a];
            indices[a] = indices[b];
            indices[b] = tmp;
        }
    }

    // Apply shuffled indices back to tiles
    state.tiles.forEach((t, i) => {
        t.frameIndex = indices[i];
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
    el.btnGenerate.disabled = false;
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

function markRegenerateStatus() {
    // Check if range or fps changed since last extraction
    if (!state.extractedWith) return;
    const current = {
        rangeIn: state.range.in,
        rangeOut: state.range.out,
        frameCount: computeSourceFrameCount()
    };
    const prev = state.extractedWith;
    const changed = current.rangeIn !== prev.rangeIn ||
                    current.rangeOut !== prev.rangeOut ||
                    current.frameCount !== prev.frameCount;
    el.btnGenerate.classList.toggle('warn', changed);
}

/* ==========================================================
   RAM ESTIMATE
   ========================================================== */

function updateRamEstimate() {
    const res = el.selectRes.value;
    const fps = parseInt(el.selectFps.value, 10);
    const dur = parseFloat(el.inputDuration.value) || 1;
    const totalFrames = Math.ceil(dur * fps);
    const mbPerFrame = RAM_PER_FRAME_BY_RES[res] || 0.3;
    const totalMB = totalFrames * mbPerFrame;

    el.ramLabel.innerText = `${totalFrames} frames @ ${res}`;
    el.ramValue.innerText = totalMB < 100 ? `${totalMB.toFixed(0)} MB` : `${(totalMB / 1024).toFixed(1)} GB`;

    const pct = Math.min(100, (totalMB / 512) * 100);
    el.ramFill.style.width = pct + '%';

    el.ramEstimate.classList.remove('warn', 'danger');
    if (totalMB > 1024) el.ramEstimate.classList.add('danger');
    else if (totalMB > 512) el.ramEstimate.classList.add('warn');
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
    if (state.isExporting) return;

    // Stop preview before export
    pause();

    state.isExporting = true;
    updateStatus();

    const path = getExportPath();
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
        const dims = { '720': [1280, 720], '1080': [1920, 1080], '4k': [3840, 2160] };
        [expW, expH] = dims[res] || [1920, 1080];
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

    // Setup VideoEncoder
    let frameCount = 0;
    const encoder = new VideoEncoder({
        output: (chunk, meta) => muxer.addChunk(chunk),
        error: (err) => { throw new Error(`VideoEncoder error: ${err.message}`); }
    });

    const encoderConfig = {
        codec: 'avc1.42001e', // H.264 Baseline@L3.1
        width: expW,
        height: expH,
        bitrate: bitrate,
        framerate: fps
    };

    encoder.configure(encoderConfig);

    // Render each frame
    for (let i = 0; i < totalFrames; i++) {
        // Backpressure: wait if queue is too large
        while (encoder.encodeQueueSize > 10) {
            await new Promise(r => setTimeout(r, 10));
        }

        // Safety: if encoder was closed (e.g. by an error), abort
        if (encoder.state === 'closed') {
            throw new Error('VideoEncoder closed unexpectedly during export');
        }

        // Render the frame at output frame i
        renderExportFrame(expCtx, expCanvas, i, totalFrames, fps);

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
function renderExportFrame(expCtx, expCanvas, outputFrame, totalFrames, fps) {
    expCtx.clearRect(0, 0, expCanvas.width, expCanvas.height);

    const cols = parseInt(el.inputCols.value);
    const N = state.totalFrames;
    const rangeDur = Math.max(0.1, state.range.out - state.range.in);
    const rate = N / rangeDur;
    const elapsed = outputFrame / fps; // time in seconds at this output frame

    // Scale tile positions from preview canvas to export canvas
    const scaleX = expCanvas.width / canvas.width;
    const scaleY = expCanvas.height / canvas.height;

    for (let i = 0; i < state.tiles.length; i++) {
        const tile = state.tiles[i];

        // Compute frame index via pipeline
        const fi = computeFrameIndex(
            tile, outputFrame, cols, N, rate, elapsed,
            anim.tileOffsets, anim.tilePhases, i
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
        const dims = { '720': [1280, 720], '1080': [1920, 1080], '4k': [3840, 2160] };
        [expW, expH] = dims[res] || [1920, 1080];
    }
    const even = ensureEvenDimensions(expW, expH);
    expW = even.width;
    expH = even.height;

    // Create export canvas
    const expCanvas = document.createElement('canvas');
    expCanvas.width = expW;
    expCanvas.height = expH;
    const expCtx = expCanvas.getContext('2d');

    // Setup MediaRecorder
    const stream = expCanvas.captureStream(0);

    // Prefer VP9, fallback to VP8
    let mimeType = 'video/webm;codecs=vp9';
    if (!MediaRecorder.isTypeSupported(mimeType)) {
        mimeType = 'video/webm;codecs=vp8';
        if (!MediaRecorder.isTypeSupported(mimeType)) {
            mimeType = 'video/webm';
        }
    }

    const recorder = new MediaRecorder(stream, { mimeType });
    const chunks = [];

    recorder.ondataavailable = (e) => {
        if (e.data.size > 0) chunks.push(e.data);
    };

    recorder.start(100); // timeslice 100ms

    // Render frames with requestAnimationFrame sync
    // captureStream(0) captures the canvas at compositor refresh rate.
    // We render each frame then wait for the next rAF to ensure the
    // compositor has picked up the new canvas content.
    for (let i = 0; i < totalFrames; i++) {
        renderExportFrame(expCtx, expCanvas, i, totalFrames, fps);

        // Wait for next frame using requestAnimationFrame to sync with compositor
        await waitFrames(1);

        el.progressFill.style.width = `${((i + 1) / totalFrames) * 100}%`;
    }

    // Flush: wait 300ms before stopping
    await new Promise(r => setTimeout(r, 300));

    // Set onstop BEFORE calling stop() to avoid race condition
    const stopPromise = new Promise(r => { recorder.onstop = r; });
    recorder.stop();
    await stopPromise;

    const blob = new Blob(chunks, { type: mimeType });

    // Generate filename
    const resLabel = res === 'preview' ? `${canvas.width}x${canvas.height}` : res;
    const ts = Date.now();
    const filename = `PanoTile-${resLabel}-${fps}fps-${dur}s-${ts}.webm`;

    downloadBlob(blob, filename);

    showToast(`Exported ${filename} (${formatBytes(blob.size)})`, 'info');
}

/**
 * Wait for N animation frames.
 */
function waitFrames(n) {
    return new Promise(resolve => {
        let count = 0;
        const cb = () => {
            count++;
            if (count >= n) resolve();
            else requestAnimationFrame(cb);
        };
        requestAnimationFrame(cb);
    });
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
        const dims = { '720': [1280, 720], '1080': [1920, 1080], '4k': [3840, 2160] };
        [expW, expH] = dims[res] || [1920, 1080];
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
el.btnGenerate.addEventListener('click', async () => {
    if (!state.ready) return;
    const ok = await confirm(
        'Regenerate grid?',
        'This will reset all pinned tiles and time positions.',
        'Regenerate'
    );
    if (ok) {
        pause();
        initProject();
        renderAll();
    }
});

// Export
el.btnExportVideo.addEventListener('click', exportVideo);
el.btnExportPng.addEventListener('click', exportPng);

// Duration slider
el.inputDuration.addEventListener('input', () => {
    el.durationValue.innerText = parseFloat(el.inputDuration.value).toFixed(1);
    updateRamEstimate();
});

// Resolution / FPS
el.selectRes.addEventListener('change', updateRamEstimate);
el.selectFps.addEventListener('change', updateRamEstimate);

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

// Square tiles
el.checkSquare.addEventListener('change', () => {
    if (state.ready) {
        pause();
        initProject();
        renderAll();
    }
});

// Cols/Rows change
el.inputCols.addEventListener('change', () => {
    if (state.ready) {
        pause();
        initProject();
        renderAll();
    }
});
el.inputRows.addEventListener('change', () => {
    if (state.ready && !el.checkSquare.checked) {
        pause();
        initProject();
        renderAll();
    }
});

/* ==========================================================
   KEYBOARD SHORTCUTS
   ========================================================== */

document.addEventListener('keydown', (e) => {
    // Ignore if user is typing in an input
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;

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
updateRamEstimate();

// Tutorial: auto-show on first visit
if (!localStorage.getItem('panotile_tutorial_seen')) {
    localStorage.setItem('panotile_tutorial_seen', '1');
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

