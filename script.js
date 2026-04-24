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
    ready: false
};

const anim = {
    mode: 'static',
    raf: null,
    lastNow: 0,
    elapsed: 0,
    tileOffsets: [],
    tilePhases: [],
    _loop: null,
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
    shuffleAmtRow: $('shuffleAmtRow'),
    inputShuffleAmt: $('inputShuffleAmt'),
    shuffleAmtValue: $('shuffleAmtValue'),
    selectPattern: $('selectPattern'),
    patternAmtRow: $('patternAmtRow'),
    inputPatternAmt: $('inputPatternAmt'),
    patternAmtValue: $('patternAmtValue'),
    btnPatternApply: $('btnPatternApply'),
    btnPatternClear: $('btnPatternClear'),
    btnBrowse: $('btnBrowse'),
    emptyState: $('emptyState'),
    loadingOverlay: $('loadingOverlay'),
    loadingText: $('loadingText').parentElement ? $('loadingText') : null,
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
    confirmCancel: $('confirmCancel')
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
    updateStatus();
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
    if (anim.mode !== 'static') return;
    const t = tileAtEvent(e);
    if (t && !t.isPinned) {
        state.activeTile = t;
        state.hoverTile = t;
        state.isDragging = true;
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

    if (anim.mode === 'spatial' && mode !== 'spatial') {
        state.tiles.forEach(t => { t.srcX = t.origSrcX; t.srcY = t.origSrcY; });
    }

    anim.mode = mode;
    if (el.shuffleAmtRow) el.shuffleAmtRow.style.display = mode === 'spatial' ? '' : 'none';

    if (mode === 'static') {
        canvas.style.cursor = 'ew-resize';
        renderAll();
        return;
    }

    canvas.style.cursor = 'default';

    if (mode === 'spatial') {
        el.inputShuffleAmt.disabled = false;
        if (state.ready) applySpatialShuffle(parseInt(el.inputShuffleAmt.value));
        return;
    }
    el.inputShuffleAmt.disabled = true;

    if (!state.ready) return;

    const n = state.tiles.length;
    const N = state.totalFrames;
    const rng = _seededRng(Date.now() | 0);
    anim.tileOffsets = Array.from({length: n}, () => rng());
    anim.tilePhases = anim.tileOffsets.map(r => r * N);
    anim.elapsed = 0;
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
}

function tickAnim(dt) {
    const cols = parseInt(el.inputCols.value);
    const N = state.totalFrames;
    const rangeDur = Math.max(0.1, state.range.out - state.range.in);
    const rate = N / rangeDur;
    const t = anim.elapsed;

    switch (anim.mode) {
        case 'linear-lr': {
            const phase = (t * rate) % N;
            state.tiles.forEach((tile, i) => {
                if (tile.isPinned) return;
                const col = i % cols;
                tile.frameIndex = Math.round((phase + (col / Math.max(1, cols - 1)) * N * 0.6) % N);
            });
            break;
        }
        case 'linear-rl': {
            const phase = (t * rate) % N;
            state.tiles.forEach((tile, i) => {
                if (tile.isPinned) return;
                const col = i % cols;
                tile.frameIndex = Math.round((phase + ((cols - 1 - col) / Math.max(1, cols - 1)) * N * 0.6) % N);
            });
            break;
        }
        case 'temporal-shuffle': {
            const phase = (t * rate) % N;
            state.tiles.forEach((tile, i) => {
                if (tile.isPinned) return;
                tile.frameIndex = Math.round((phase + anim.tileOffsets[i] * N) % N);
            });
            break;
        }
        case 'drunk': {
            state.tiles.forEach((tile, i) => {
                if (tile.isPinned) return;
                const col = i % cols;
                const row = Math.floor(i / cols);
                const v = valueNoise3(
                    col * 0.7 + anim.tileOffsets[i] * 5.1,
                    row * 0.7 + anim.tileOffsets[i] * 3.7,
                    t * 0.6
                );
                tile.frameIndex = Math.round(v * (N - 1));
            });
            break;
        }
        case 'perlin-flow': {
            state.tiles.forEach((tile, i) => {
                if (tile.isPinned) return;
                const col = i % cols;
                const row = Math.floor(i / cols);
                const v = valueNoise3(col * 0.35, row * 0.35, t * 0.12) * 2 - 1;
                anim.tilePhases[i] = ((anim.tilePhases[i] + v * rate * dt) % N + N) % N;
                tile.frameIndex = Math.round(anim.tilePhases[i]);
            });
            break;
        }
    }
}

function applySpatialShuffle(amount) {
    const n = state.tiles.length;
    const perm = Array.from({length: n}, (_, i) => i);
    const swaps = Math.round((amount / 10) * n);
    const rng = _seededRng(12345);
    for (let i = 0; i < swaps; i++) {
        const j = i + Math.floor(rng() * (n - i));
        [perm[i], perm[j]] = [perm[j], perm[i]];
    }
    state.tiles.forEach((t, i) => {
        const src = state.tiles[perm[i]];
        t.srcX = src.origSrcX;
        t.srcY = src.origSrcY;
    });
    renderAll();
}

el.selectMode.addEventListener('change', () => setMode(el.selectMode.value));

el.inputShuffleAmt.addEventListener('input', () => {
    el.shuffleAmtValue.innerText = el.inputShuffleAmt.value;
    if (anim.mode === 'spatial') applySpatialShuffle(parseInt(el.inputShuffleAmt.value));
});

/* ==========================================================
   EXPORT — PNG
   ========================================================== */

el.btnExportPng.onclick = () => {
    state.isExporting = true;     // hides pins + grid
    renderAll();
    try {
        const a = document.createElement('a');
        a.download = `PanoTile-${Date.now()}.png`;
        a.href = canvas.toDataURL('image/png', 1.0);
        a.click();
        showToast('PNG exported', 'info');
    } catch (err) {
        console.error(err);
        showToast('PNG export failed.');
    } finally {
        state.isExporting = false;
        renderAll();
    }
};

/* ==========================================================
   EXPORT — VIDEO (WebM) with resolution + FPS selectors
   ========================================================== */

el.btnExportVideo.onclick = exportVideo;

async function exportVideo() {
    const fps = parseInt(el.selectFps.value, 10);
    const resMode = el.selectRes.value;
    const isLoop = el.checkLoop.checked;

    // Compute export dimensions
    const vRatio = state.video.width / state.video.height;
    let targetW, targetH;
    switch (resMode) {
        case '720':  targetH = 720;  targetW = Math.round(720 * vRatio); break;
        case '1080': targetH = 1080; targetW = Math.round(1080 * vRatio); break;
        case '4k':   targetH = 2160; targetW = Math.round(2160 * vRatio); break;
        default:     targetW = canvas.width; targetH = canvas.height;
    }
    // H.264 requires even dimensions
    targetW = targetW - (targetW % 2);
    targetH = targetH - (targetH % 2);

    // Off-screen canvas at target res
    const cols = parseInt(el.inputCols.value);
    const rows = parseInt(el.inputRows.value);
    const exportCanvas = document.createElement('canvas');
    exportCanvas.width = targetW;
    exportCanvas.height = targetH;
    const exCtx = exportCanvas.getContext('2d');

    // Recompute crop
    const cRatio = targetW / targetH;
    let crW = state.video.width, crH = state.video.height;
    if (cRatio > vRatio) crH = state.video.width / cRatio;
    else crW = state.video.height * cRatio;
    const sW = crW / cols, sH = crH / rows;
    const oX = (state.video.width - crW) / 2;
    const oY = (state.video.height - crH) / 2;

    const exTiles = state.tiles.map((t, i) => {
        const c = i % cols, r = Math.floor(i / cols);
        return {
            x: c * (targetW / cols), y: r * (targetH / rows),
            w: targetW / cols, h: targetH / rows,
            srcX: oX + c * sW, srcY: oY + r * sH,
            srcW: sW, srcH: sH,
            frameIndex: t.frameIndex, isPinned: t.isPinned
        };
    });

    const exportDuration = parseFloat(el.inputDuration.value);
    const totalOutputFrames = Math.round(exportDuration * fps);

    const animLoopSaved = anim._loop;
    if (anim.raf) { cancelAnimationFrame(anim.raf); anim.raf = null; }

    state.isExporting = true;
    updateStatus();

    // Decide which export path to use.
    // WebCodecs is PREFERRED: deterministic timestamps, real MP4/H.264,
    // exact fps, exact duration. No realtime dependency.
    const hasWebCodecs = typeof VideoEncoder !== 'undefined'
                      && typeof Mp4Muxer !== 'undefined'
                      && typeof Mp4Muxer.Muxer === 'function';

    try {
        if (hasWebCodecs) {
            await exportMp4WebCodecs({
                exportCanvas, exCtx, exTiles, targetW, targetH,
                fps, totalOutputFrames, isLoop, resMode, exportDuration
            });
        } else {
            console.warn('WebCodecs unavailable — falling back to MediaRecorder WebM');
            await exportWebmMediaRecorder({
                exportCanvas, exCtx, exTiles, targetW, targetH,
                fps, totalOutputFrames, isLoop, resMode, exportDuration
            });
        }
    } catch (err) {
        console.error('Export failed:', err);
        showToast(`Export failed: ${err.message || err}`);
    } finally {
        state.isExporting = false;
        hideLoading();
        renderAll();
        updateStatus();
        if (animLoopSaved && anim.mode !== 'static' && anim.mode !== 'spatial') {
            anim.lastNow = performance.now();
            anim.raf = requestAnimationFrame(animLoopSaved);
        }
    }
}

/* ----------------------------------------------------------
   PATH A — WebCodecs → MP4 (deterministic)
   Every frame gets an exact timestamp. Encoder output goes
   straight to mp4-muxer. File duration = exact.
   ---------------------------------------------------------- */
async function exportMp4WebCodecs({
    exportCanvas, exCtx, exTiles, targetW, targetH,
    fps, totalOutputFrames, isLoop, resMode, exportDuration
}) {
    showLoading('ENCODING MP4');

    const bitrate = { '720': 6e6, '1080': 10e6, '4k': 20e6, 'preview': 4e6 }[resMode] || 10e6;

    // Pick an H.264 profile. 'avc1.640028' = High@L4, good for 1080p.
    // For 4K we'd want L5.1 but most browsers auto-adjust.
    const codecString = 'avc1.42E01F'; // Baseline@L3.1 — widest compatibility

    // Check encoder support before building anything
    const support = await VideoEncoder.isConfigSupported({
        codec: codecString,
        width: targetW,
        height: targetH,
        bitrate,
        framerate: fps
    });
    if (!support.supported) {
        throw new Error('H.264 encoder rejected this config');
    }

    // Build the muxer
    const muxer = new Mp4Muxer.Muxer({
        target: new Mp4Muxer.ArrayBufferTarget(),
        video: {
            codec: 'avc',
            width: targetW,
            height: targetH,
            frameRate: fps
        },
        fastStart: 'in-memory',   // moov at the start — playable everywhere
        firstTimestampBehavior: 'offset'
    });

    // Encoder that pushes each chunk into the muxer
    const encoder = new VideoEncoder({
        output: (chunk, meta) => muxer.addVideoChunk(chunk, meta),
        error: (e) => { throw e; }
    });
    encoder.configure({
        codec: codecString,
        width: targetW,
        height: targetH,
        bitrate,
        framerate: fps,
        // Every 2 seconds force a keyframe (helps with seeking)
        // Handled manually below for precision
    });

    const keyframeInterval = fps * 2;  // keyframe every 2s

    for (let i = 0; i < totalOutputFrames; i++) {
        // Advance tiles
        if (i > 0) {
            exTiles.forEach(t => {
                if (t.isPinned) return;
                t.frameIndex++;
                if (t.frameIndex >= state.totalFrames) {
                    t.frameIndex = isLoop ? 0 : state.totalFrames - 1;
                }
            });
        }
        drawExportFrame(exCtx, exTiles, targetW, targetH);

        // Create a VideoFrame from the canvas with a precise timestamp.
        // Microseconds. This is the source of truth — muxer uses it directly.
        const timestamp = Math.round((i * 1_000_000) / fps);
        const duration = Math.round(1_000_000 / fps);

        const videoFrame = new VideoFrame(exportCanvas, {
            timestamp,
            duration
        });

        const keyFrame = (i % keyframeInterval === 0);
        encoder.encode(videoFrame, { keyFrame });
        videoFrame.close();  // release GPU resource

        el.progressFill.style.width = `${((i + 1) / totalOutputFrames) * 100}%`;

        // Backpressure: if the encoder queue gets long, wait.
        // This keeps memory bounded on long exports.
        if (encoder.encodeQueueSize > 10) {
            while (encoder.encodeQueueSize > 4) {
                await sleep(10);
            }
        }
    }

    await encoder.flush();
    encoder.close();
    muxer.finalize();

    const buffer = muxer.target.buffer;
    const blob = new Blob([buffer], { type: 'video/mp4' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `PanoTile-${resMode}-${fps}fps-${exportDuration}s-${Date.now()}.mp4`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 2000);

    showToast('MP4 exported', 'info');
}

/* ----------------------------------------------------------
   PATH B — MediaRecorder → WebM (fallback)
   For browsers without WebCodecs (Safari <16.4, Firefox).
   Best-effort: output may have mild timing drift.
   ---------------------------------------------------------- */
async function exportWebmMediaRecorder({
    exportCanvas, exCtx, exTiles, targetW, targetH,
    fps, totalOutputFrames, isLoop, resMode, exportDuration
}) {
    if (typeof MediaRecorder === 'undefined') {
        throw new Error('Neither WebCodecs nor MediaRecorder available');
    }
    showLoading('RENDERING WEBM (fallback)');

    let options = {};
    const bitrate = { '720': 5e6, '1080': 8e6, '4k': 15e6, 'preview': 3e6 }[resMode] || 8e6;
    options.videoBitsPerSecond = bitrate;
    if (MediaRecorder.isTypeSupported('video/webm;codecs=vp9')) options.mimeType = 'video/webm;codecs=vp9';
    else if (MediaRecorder.isTypeSupported('video/webm;codecs=vp8')) options.mimeType = 'video/webm;codecs=vp8';
    else if (MediaRecorder.isTypeSupported('video/webm')) options.mimeType = 'video/webm';
    else if (MediaRecorder.isTypeSupported('video/mp4')) options.mimeType = 'video/mp4';

    const stream = exportCanvas.captureStream(0);
    const videoTrack = stream.getVideoTracks()[0];
    const hasRequestFrame = !!videoTrack.requestFrame;
    if (!hasRequestFrame) {
        // Truly old browsers: fall back to timed capture
        stream.getVideoTracks().forEach(t => t.stop());
        const timedStream = exportCanvas.captureStream(fps);
        return exportWebmTimed(timedStream, options, {
            exportCanvas, exCtx, exTiles, targetW, targetH,
            fps, totalOutputFrames, isLoop, resMode, exportDuration
        });
    }

    const recorder = new MediaRecorder(stream, options);
    const chunks = [];
    recorder.ondataavailable = (e) => { if (e.data.size > 0) chunks.push(e.data); };

    return new Promise((resolve) => {
        recorder.onstop = () => {
            const ext = options.mimeType && options.mimeType.includes('mp4') ? 'mp4' : 'webm';
            const blob = new Blob(chunks, { type: options.mimeType || 'video/webm' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `PanoTile-${resMode}-${fps}fps-${exportDuration}s-${Date.now()}.${ext}`;
            a.click();
            setTimeout(() => URL.revokeObjectURL(url), 2000);
            showToast('Video exported (WebM)', 'info');
            resolve();
        };

        (async () => {
            drawExportFrame(exCtx, exTiles, targetW, targetH);
            recorder.start(100);
            await waitFrames(3);
            videoTrack.requestFrame();
            await waitFrames(2);

            for (let i = 0; i < totalOutputFrames; i++) {
                if (i > 0) {
                    exTiles.forEach(t => {
                        if (t.isPinned) return;
                        t.frameIndex++;
                        if (t.frameIndex >= state.totalFrames) {
                            t.frameIndex = isLoop ? 0 : state.totalFrames - 1;
                        }
                    });
                    drawExportFrame(exCtx, exTiles, targetW, targetH);
                }
                await waitFrames(1);
                videoTrack.requestFrame();
                const settle = (resMode === '4k') ? 3 : (resMode === '1080' ? 2 : 1);
                await waitFrames(settle);
                el.progressFill.style.width = `${((i + 1) / totalOutputFrames) * 100}%`;
            }
            await sleep(300);
            recorder.stop();
        })();
    });
}

// Very old browsers path — sleep-paced capture. Not frame-perfect.
async function exportWebmTimed(stream, options, params) {
    const { exportCanvas, exCtx, exTiles, targetW, targetH,
            fps, totalOutputFrames, isLoop, resMode, exportDuration } = params;
    const recorder = new MediaRecorder(stream, options);
    const chunks = [];
    recorder.ondataavailable = (e) => { if (e.data.size > 0) chunks.push(e.data); };
    return new Promise((resolve) => {
        recorder.onstop = () => {
            const ext = options.mimeType && options.mimeType.includes('mp4') ? 'mp4' : 'webm';
            const blob = new Blob(chunks, { type: options.mimeType || 'video/webm' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `PanoTile-${resMode}-${fps}fps-${exportDuration}s-${Date.now()}.${ext}`;
            a.click();
            setTimeout(() => URL.revokeObjectURL(url), 2000);
            showToast('Video exported (timed fallback)', 'info');
            resolve();
        };
        (async () => {
            drawExportFrame(exCtx, exTiles, targetW, targetH);
            recorder.start(100);
            for (let i = 0; i < totalOutputFrames; i++) {
                if (i > 0) {
                    exTiles.forEach(t => {
                        if (t.isPinned) return;
                        t.frameIndex++;
                        if (t.frameIndex >= state.totalFrames) {
                            t.frameIndex = isLoop ? 0 : state.totalFrames - 1;
                        }
                    });
                    drawExportFrame(exCtx, exTiles, targetW, targetH);
                }
                await sleep(1000 / fps);
                el.progressFill.style.width = `${((i + 1) / totalOutputFrames) * 100}%`;
            }
            await sleep(300);
            recorder.stop();
        })();
    });
}

function waitFrames(n) {
    return new Promise(resolve => {
        let remaining = n;
        const tick = () => {
            if (--remaining <= 0) resolve();
            else requestAnimationFrame(tick);
        };
        requestAnimationFrame(tick);
    });
}

function sleep(ms) {
    return new Promise(r => setTimeout(r, ms));
}

function drawExportFrame(exCtx, exTiles, w, h) {
    exCtx.clearRect(0, 0, w, h);
    for (const t of exTiles) {
        const img = state.frames[t.frameIndex];
        if (img) {
            exCtx.drawImage(img, t.srcX, t.srcY, t.srcW, t.srcH, t.x, t.y, t.w, t.h);
        }
    }
}

/* ==========================================================
   LOADING
   ========================================================== */

function showLoading(text) {
    $('loadingText').innerText = text;
    el.progressFill.style.width = '0%';
    el.loadingOverlay.classList.add('visible');
}
function hideLoading() {
    el.loadingOverlay.classList.remove('visible');
}

/* ==========================================================
   CONTROL ENABLEMENT
   ========================================================== */

function enableControls() {
    const ctrls = [
        el.inputCols, el.inputRows, el.checkSquare, el.checkGrid, el.checkLoop,
        el.selectRes, el.selectFps, el.inputDuration, el.selectMode,
        el.btnGenerate, el.btnExportPng, el.btnExportVideo,
        el.selectPattern, el.btnPatternClear
    ];
    ctrls.forEach(c => c.disabled = false);
    if (el.checkSquare.checked) el.inputRows.disabled = true;
}

/* ==========================================================
   MEMORY ESTIMATOR
   ========================================================== */

function updateRamEstimate() {
    const frames = computeSourceFrameCount();
    const resMode = el.selectRes.value;
    const mbPerFrame = RAM_PER_FRAME_BY_RES[resMode] || 0.3;
    const totalMB = frames * mbPerFrame;

    // Soft budgets based on typical browser memory headroom
    const WARN_MB = 300;
    const DANGER_MB = 600;

    el.ramLabel.innerText = `${frames} frames`;
    el.ramValue.innerText = `~${totalMB.toFixed(0)} MB`;

    const pct = Math.min(100, (totalMB / DANGER_MB) * 100);
    el.ramFill.style.width = pct + '%';

    el.ramEstimate.classList.remove('warn', 'danger');
    if (totalMB >= DANGER_MB) el.ramEstimate.classList.add('danger');
    else if (totalMB >= WARN_MB) el.ramEstimate.classList.add('warn');
}

function markRegenerateStatus() {
    if (!state.extractedWith) {
        el.btnGenerate.classList.remove('btn-primary');
        return;
    }
    const ew = state.extractedWith;
    const currentCount = computeSourceFrameCount();
    const stale =
        ew.rangeIn !== state.range.in ||
        ew.rangeOut !== state.range.out ||
        ew.frameCount !== currentCount;
    el.btnGenerate.classList.toggle('btn-primary', stale);
    el.btnGenerate.innerText = stale ? 'Regenerate (stale)' : 'Regenerate grid';
}

/* ==========================================================
   CONTROL HANDLERS
   ========================================================== */

el.btnGenerate.onclick = async () => {
    if (!state.ready) return;

    // Detect if re-extraction is needed (duration or fps changed → different frame count)
    const newFrameCount = computeSourceFrameCount();
    const needsReextract = newFrameCount !== state.totalFrames;

    const hasWork = state.tiles.some(t => t.isPinned || t.frameIndex !== 0);
    if (hasWork) {
        const msg = needsReextract
            ? 'Duration/FPS changed. This will re-extract frames and reset pins and time positions.'
            : 'This will reset pinned tiles and time positions.';
        const ok = await confirm('Regenerate grid?', msg, 'Regenerate');
        if (!ok) return;
    }

    if (needsReextract && state.video.probe) {
        extractFrames(state.video.probe);
    } else {
        initProject();
        markRegenerateStatus();
    }
};

el.inputDuration.addEventListener('input', (e) => {
    const v = parseFloat(e.target.value);
    el.durationValue.innerText = v.toFixed(1);
    // Duration no longer affects source frames — only output length
});

el.selectFps.addEventListener('change', () => { updateRamEstimate(); markRegenerateStatus(); });
el.selectRes.addEventListener('change', updateRamEstimate);

el.checkSquare.onchange = (e) => {
    el.inputRows.disabled = e.target.checked;
};

el.checkGrid.onchange = () => renderAll();

// Clamp numeric inputs on blur
[el.inputCols, el.inputRows].forEach(input => {
    input.addEventListener('blur', () => {
        let v = parseInt(input.value, 10);
        if (!isFinite(v) || v < 1) v = 1;
        if (v > 30) v = 30;
        input.value = v;
    });
});

/* ==========================================================
   BLOCK PATTERN
   ========================================================== */

function applyPattern() {
    const pattern = el.selectPattern.value;
    if (pattern === 'none') return;

    const cols = parseInt(el.inputCols.value);
    const rows = parseInt(el.inputRows.value);

    state.tiles.forEach(t => { t.isPinned = false; });

    switch (pattern) {
        case 'checkerboard':
            state.tiles.forEach((t, i) => {
                const col = i % cols, row = Math.floor(i / cols);
                t.isPinned = (row + col) % 2 === 0;
            });
            break;

        case 'random': {
            const pct = parseInt(el.inputPatternAmt.value) / 100;
            state.tiles.forEach(t => { t.isPinned = Math.random() < pct; });
            break;
        }

        case 'borders':
            state.tiles.forEach((t, i) => {
                const col = i % cols, row = Math.floor(i / cols);
                t.isPinned = row === 0 || row === rows - 1 || col === 0 || col === cols - 1;
            });
            break;

        case 'center': {
            const r0 = Math.floor(rows / 4), r1 = Math.ceil(3 * rows / 4);
            const c0 = Math.floor(cols / 4), c1 = Math.ceil(3 * cols / 4);
            state.tiles.forEach((t, i) => {
                const col = i % cols, row = Math.floor(i / cols);
                t.isPinned = row >= r0 && row < r1 && col >= c0 && col < c1;
            });
            break;
        }
    }

    renderAll();
    updateStatus();
}

function clearAllPins() {
    state.tiles.forEach(t => { t.isPinned = false; });
    renderAll();
    updateStatus();
}

el.selectPattern.addEventListener('change', () => {
    const p = el.selectPattern.value;
    el.patternAmtRow.style.display = p === 'random' ? '' : 'none';
    el.btnPatternApply.disabled = p === 'none';
    if (p === 'random') el.inputPatternAmt.disabled = false;
});

el.inputPatternAmt.addEventListener('input', () => {
    el.patternAmtValue.innerText = el.inputPatternAmt.value;
});

el.btnPatternApply.onclick = applyPattern;
el.btnPatternClear.onclick = clearAllPins;

/* ==========================================================
   KEYBOARD SHORTCUTS
   ========================================================== */

window.addEventListener('keydown', (e) => {
    if (!state.ready) return;
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
    switch (e.key.toLowerCase()) {
        case 'g': el.checkGrid.checked = !el.checkGrid.checked; renderAll(); break;
    }
});

