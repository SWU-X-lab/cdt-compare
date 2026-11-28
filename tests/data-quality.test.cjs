'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'exp1_fcdt_compare.js'), 'utf8');
const adapterSource = fs.readFileSync(path.join(root, 'src/storage/supabase-storage.js'), 'utf8');
const entries = ['index.html', 'exp1_fcdt_compare.html'];

function browser(entry = entries[0], group = '0', code = '', resume = null, options = {}) {
    const html = fs.readFileSync(path.join(root, entry), 'utf8');
    const ids = new Map();
    const calls = [];
    const downloads = [];
    const timers = new Map();
    let timerId = 0;
    let time = 0;
    const context2D = new Proxy({}, { get: () => () => {} });
    function target() {
        const listeners = new Map();
        return {
            addEventListener(name, fn) {
                const list = listeners.get(name) || [];
                list.push(fn);
                listeners.set(name, list);
            },
            removeEventListener(name, fn) {
                listeners.set(name, (listeners.get(name) || []).filter(x => x !== fn));
            },
            async emit(name, event = {}) {
                for (const fn of (listeners.get(name) || []).slice()) {
                    await fn({ target: this, preventDefault() {}, stopPropagation() {}, ...event });
                }
            }
        };
    }
    for (const match of html.matchAll(/\bid="([^"]+)"/g)) {
        const classes = new Set(match[1].includes('manual-download') ? ['hidden'] : []);
        ids.set(match[1], {
            id: match[1], ...target(), style: {},
            classList: { add(cls) { classes.add(cls); }, remove(cls) { classes.delete(cls); }, contains(cls) { return classes.has(cls); } },
            getContext: () => context2D, reportValidity: () => true, focus() {}
        });
    }
    const document = {
        ...target(), hidden: false, pointerLockElement: null, fullscreenElement: null,
        getElementById: name => ids.get(name) || null,
        createElement: () => ({style: {}, click() { downloads.push([this.href, this.download]); }, remove() {}}),
        body: { appendChild() {} },
        exitPointerLock() { this.pointerLockElement = null; },
        exitFullscreen() { this.fullscreenElement = null; }
    };
    document.documentElement = { requestFullscreen() { document.fullscreenElement = this; } };
    const canvas = ids.get('expCanvas');
    canvas.requestPointerLock = () => {
        document.pointerLockElement = canvas;
        void document.emit('pointerlockchange');
    };
    canvas.getBoundingClientRect = () => ({ width: canvas.width, height: canvas.height });
    const storage = Object.fromEntries([
        'initializeStorage', 'loadProgress', 'saveTrial', 'saveCheckpoint', 'saveProgress',
        'clearProgress', 'saveFinalResult', 'finishExperiment'
    ].map(name => [name, async (value) => {
        calls.push([name, value]);
        if (name === 'saveFinalResult' && options.cloudSave) return options.cloudSave(value);
        return name === 'saveFinalResult' ? { saved: true } : name === 'loadProgress' ? resume : null;
    }]));
    const window = { ...target(), CLTStorage: storage, innerWidth: 1024, innerHeight: 768 };
    const addTimer = (fn, delay, interval = 0) => {
        const id = ++timerId;
        timers.set(id, { at: time + delay, fn, interval });
        return id;
    };
    const scope = vm.createContext({
        window, document, HTMLInputElement: class {}, HTMLTextAreaElement: class {}, HTMLSelectElement: class {}, history: { pushState() {} },
        console: { log() {}, warn() {}, error() {} },
        alert: value => calls.push(['alert', value]),
        performance: { now: () => time },
        setTimeout: addTimer, clearTimeout: id => timers.delete(id),
        setInterval: (fn, delay) => addTimer(fn, delay, delay), clearInterval: id => timers.delete(id),
        requestAnimationFrame: fn => addTimer(() => fn(time), 50),
        Uint8Array, Blob, URL: { createObjectURL: () => 'blob:test-xlsx' },
        XLSX: {
            utils: {
                json_to_sheet: rows => rows,
                book_new: () => ({SheetNames: []}),
                book_append_sheet: (book, _sheet, name) => book.SheetNames.push(name)
            },
            write: () => {
                if (options.localFails) throw new Error('模拟 Excel 失败');
                return Uint8Array.of(0x50, 0x4B);
            }
        },
        exportsForTest: []
    });
    vm.runInContext(source, scope, { filename: 'exp1_fcdt_compare.js' });
    const evaluate = expression => vm.runInContext(expression, scope);
    if (!options.realExport) evaluate('downloadXLSX = () => exportsForTest.push(JSON.parse(JSON.stringify(stimData)))');
    ids.get('subGroup').value = group;
    ids.get('participantCode').value = code;
    async function until(expression) {
        for (let i = 0; i < 5000; i++) {
            for (let j = 0; j < 8; j++) await Promise.resolve();
            if (evaluate(expression)) return;
            const pending = [...timers.entries()].sort((a, b) => a[1].at - b[1].at)[0];
            if (!pending) break;
            const [id, t] = pending;
            time = t.at;
            if (t.interval) t.at += t.interval;
            else timers.delete(id);
            t.fn();
            for (let j = 0; j < 5; j++) await Promise.resolve();
        }
        assert.fail('未达到预期流程状态：' + expression + '，当前：' + evaluate('experimentPhase'));
    }
    const space = () => window.emit('keydown', { code: 'Space', key: ' ' });
    return { evaluate, until, space, calls, downloads, exportsForTest: scope.exportsForTest,
        start: () => ids.get('start-btn').emit('click'), canvas, document, window };
}

test('条件缺失在 Excel 留空，真实零保留，超时和中断分别计数', () => {
    const { evaluate } = browser();
    const x = evaluate("kSheetFromRows([buildKRow(4,250,[{isChange:1,accuracy:1,CountedInK:1},{isChange:0,timeout:true,CountedInK:0},{isChange:1,TrialStatus:'interrupted',CountedInK:0}])])[0]");
    assert.equal(x.HitRate, 1);
    assert.equal(Object.hasOwn(x, 'FalseAlarmRate'), false);
    assert.equal(Object.hasOwn(x, 'K'), false);
    assert.equal(x.TimeoutTrials, 1);
    assert.equal(x.InterruptedTrials, 1);
    const z = evaluate("kSheetFromRows([buildKRow(4,250,[{isChange:1,accuracy:0},{isChange:0,response:0,accuracy:1}])])[0]");
    assert.equal(z.K, 0);
    assert.equal(z.FalseAlarmRate, 0);
});

test('圆环布局的每个色块都不跨左右中线，两边都可选', () => {
    const { evaluate } = browser();
    const ok = evaluate("(() => { for (const [w,h] of [[320,240],[1024,768],[768,1024],[1920,1080]]) { canvas.width=w;canvas.height=h; for (const size of [4,6]) for (let i=0;i<300;i++) { const locs=getStimLocs(size,w/2,h/2); if (locs.length!==size || !locs.some(p=>p.x<w/2) || !locs.some(p=>p.x>w/2) || locs.some(p=>Math.abs(p.x-w/2)<26.49)) return false; } } return true; })()");
    assert.equal(ok, true);
});

for (const entry of entries) {
    test(entry + '：调试组从开始到本地导出，不调用任何远端方法', async () => {
        const b = browser(entry);
        await b.start();
        assert.equal(b.evaluate('experimentPhase'), 'instructions');
        await b.space();
        await b.until("experimentPhase === 'practice_end'");
        await b.space();
        await b.until("experimentPhase === 'ending_wait'");
        await b.space();
        await b.until("experimentPhase === 'finished'");
        assert.deepEqual(b.calls, []);
        assert.equal(b.exportsForTest[0].length, 5);
        assert.ok(b.exportsForTest[0].every(row => row.TrialStatus === 'timeout'));
    });
}

test('正式组中断：保留一条中断记录，取消超时、不上传，并且不从头编号', async () => {
    const b = browser(entries[0], '1', 'Test_A17', { completedBlocks: 1 });
    await b.start();
    assert.equal(b.evaluate('p.subjectID'), 'Test_A17');
    assert.equal(b.evaluate('p.rndSeed'), b.evaluate("seedFromSubjectID('Test_A17')"));
    assert.notEqual(b.evaluate('p.rndSeed'), b.evaluate("seedFromSubjectID('Test_A18')"));
    await b.space();
    await b.until("experimentPhase === 'trial_response'");
    b.document.pointerLockElement = null;
    await b.document.emit('pointerlockchange');
    await b.until("experimentPhase === 'finished'");
    assert.deepEqual(b.calls.map(([name]) => name), ['initializeStorage','loadProgress']);
    assert.equal(b.exportsForTest[0].length, 1);
    const row = b.exportsForTest[0][0];
    assert.equal(row.TrialGlobal, 21);
    assert.equal(row.TrialStatus, 'interrupted');
    assert.equal(row.interruptionReason, 'pointer_lock_lost');
    assert.equal(row.CountedInK, 0);
    assert.equal(row.timeout, false);
});

test('独立窗口失焦：标签页仍可见、鼠标仍锁定时记为中断而非超时', async () => {
    const b = browser(entries[0], '1', 'Test_A17', { completedBlocks: 1 });
    await b.start();
    await b.space();
    await b.until("experimentPhase === 'trial_response'");
    assert.equal(b.document.hidden, false);
    assert.equal(b.document.pointerLockElement, b.canvas);
    assert.ok(b.document.fullscreenElement);
    await b.window.emit('blur');
    await b.until("stimData.length > 0");
    const row = b.evaluate('stimData[0]');
    assert.equal(row.TrialGlobal, 21);
    assert.equal(row.TrialStatus, 'interrupted');
    assert.equal(row.interruptionReason, 'window_blur');
    assert.equal(row.timeout, false);
    await b.until("experimentPhase === 'finished'");
    assert.deepEqual(b.calls.map(([name]) => name), ['initializeStorage', 'loadProgress']);
    assert.equal(b.exportsForTest[0].length, 1);
});

test('正式组完成第四组可确认本会话数据上传，鼠标落点字段保留', async () => {
    const b = browser(entries[0], '1', 'Test_A17', { completedBlocks: 3 });
    await b.start();
    await b.space();
    await b.until("experimentPhase === 'trial_response'");
    b.evaluate("Object.assign(virtualCursor, probeSnapshot.locs[0]); cursorTravelPx=123;");
    await b.canvas.emit('mousedown', { button: 0 });
    await b.until("experimentPhase === 'ending_wait'");
    await b.space();
    await b.until("experimentPhase === 'finished'");
    const final = b.calls.find(([name]) => name === 'saveFinalResult')[1];
    assert.equal(final.trials.length, 20);
    assert.equal(final.trials[0].TrialGlobal, 61);
    assert.equal(final.trials[0].cursorTravelPx, 123);
    assert.ok(['left','right'].includes(final.trials[0].clickSide));
    assert.equal(b.calls.at(-1)[1].successful, true);
    assert.equal(b.exportsForTest[0][0].TrialGlobal, 61);
});

test('完整 Supabase 写入失败时不删字段重试', async () => {
    const calls = [];
    const window = { supabase: { createClient: () => ({
        from() { return { async insert(rows) { calls.push(rows); return { error: { message: '字段缺失' } }; } }; }
    }) } };
    vm.runInNewContext(adapterSource, { window, console });
    const adapter = window.CLTStorageFactories.supabase();
    await adapter.initialize();
    const rows = [{ clickSide:'left', cursorTravelPx:123 }];
    await assert.rejects(adapter.saveFinalResult({ trials: rows }), /字段缺失/);
    assert.deepEqual(calls, [rows]);
});

for (const entry of entries) {
    test(entry + '：正式写入尚未返回时，Excel 已自动触发且手动链接可用', async () => {
        let rejectCloud;
        const cloud = new Promise((_resolve, reject) => { rejectCloud = reject; });
        const b = browser(entry, '1', 'Test_A17', { completedBlocks: 3 },
            { realExport: true, cloudSave: () => cloud });
        await b.start();
        await b.space();
        await b.until("experimentPhase === 'ending_wait'");
        await b.space();
        assert.equal(b.evaluate('experimentPhase'), 'upload');
        assert.deepEqual(b.downloads, [['blob:test-xlsx', 'exp1_fcdt_compare_Test_A17.xlsx']]);
        const pendingLink = b.document.getElementById('upload-manual-download-link');
        assert.equal(pendingLink.href, 'blob:test-xlsx');
        assert.equal(b.document.getElementById('upload-manual-download').classList.contains('hidden'), false);
        assert.match(b.document.getElementById('upload-status').innerText, /本地 Excel：已生成.*\n服务器：正在保存/);
        assert.equal(b.document.pointerLockElement, null);
        const payload = b.calls.find(([name]) => name === 'saveFinalResult')[1];
        assert.equal(payload.trials.length, 20);
        assert.equal(payload.trials[0].subjectID, 'Test_A17');
        assert.equal(payload.trials[0].TrialGlobal, 61);
        await b.space();
        assert.equal(b.downloads.length, 1);
        assert.equal(b.calls.filter(([name]) => name === 'saveFinalResult').length, 1);
        rejectCloud(new Error('模拟网络错误'));
        await b.until("experimentPhase === 'finished'");
        assert.match(b.document.getElementById('finished-text').innerText, /本地 Excel：已生成.*\n服务器：保存失败/);
        assert.equal(b.document.getElementById('manual-download-link').href, pendingLink.href);
        assert.equal(b.calls.filter(([name]) => name === 'saveFinalResult').length, 1);
    });
}

test('本地 Excel 导出失败仍独立尝试服务器，只有云端显示已保存', async () => {
    const b = browser(entries[0], '2', 'Test_A17', { completedBlocks: 3 },
        { realExport: true, localFails: true });
    await b.start();
    await b.space();
    await b.until("experimentPhase === 'ending_wait'");
    await b.space();
    await b.until("experimentPhase === 'finished'");
    assert.deepEqual(b.downloads, []);
    assert.match(b.document.getElementById('finished-text').innerText, /本地 Excel：导出失败.*\n服务器：已保存/);
    assert.equal(b.calls.filter(([name]) => name === 'saveFinalResult').length, 1);
});
