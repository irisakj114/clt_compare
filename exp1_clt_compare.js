// ============================================================
// exp1_clt.js
// 变化定位任务 (Change Localization Task, CLT)
// 基于 CLT.m (0630版) 实验规则
//
// 实验设计：
//   compare 设计：set size 4/6 两种条件各 40 trial，共 80 trial
//   正式试次整体随机后均分到 4 个 block（每 block 20 trial）
//   练习 10 trial（2 种 set size × 5），正式呈现时长固定为 250 ms
//
// 试次结构：
//   1. ITI              (500 ms)           显示注视点
//   2. 学习阵列         (50/200/500/1000 ms)      显示注视点，时长由当前试次条件决定
//   3. 保留间隔         (1000 ms)          显示注视点（注视点全程固定显示，不消失）
//   4. 测试阵列         (最多 4000 ms)      超时未反应记为 null
//
// Block 规则：
//   4 个 block，每 block 20 trial；block 间强制休息 15 秒（含倒计时），
//   倒计时结束后显示「您可以继续休息或按【SPACE】开始下一组实验」，按 SPACE 进入下一组
//   不设「低正确率自动结束」：无论正确率高低，实验都完整跑完 4 个 block
//   低正确率只做标记（详见下方「低正确率标记」段）
//
// 调试模式：组别填 0 进入调试；1 为正式。调试模式固定 SubjectID 与种子为 1111，
// 便于识别与复现。正式模式自动生成：{subGroup}_{3位点击顺序}_{3位随机后缀}。
// 人口学信息改为在 Credamo 另行收集，本程序不再采集。
//
// 低正确率标记（导出文件名前加 F，不改变实验流程）：
//   判定条件：set size = 3 或 4 时，「呈现时长 ≥ 200 ms」试次的正确率低于 50%
//   （set size 3 和 4 分别计算，任一低于 50% 即触发；set size = 6 不参与判定）
//   触发的条件行会在 K 值表中以「LowAccRed = 1」标记（标红功能受 XLSX 版本所限，见下）
//
// 导出（Excel，4 个 sheet，调试模式与正式实验规则完全一致）：
//   文件名：实验名称_被试编号（如 exp1_clt_compare_1_023_abc.xlsx）；
//           存在低正确率条件时在最前面加 F；调试模式不做该判定。
//   Sheet1「Trial Data」：全部 trial，按实验实际发生顺序排列
//   Sheet2「K Values」：条件级正确率 / 容量 K / 两种超时口径
//   Sheet3「ColorAccuracy」：每个颜色在每个 set size 下的正确率
//   Sheet4「Grouped Trials」：按「呈现时长 → set size」分组重排
//   Sheet2「K Values」：按「set size × 呈现时长」拆分的 12 个条件各一行，
//                      给出正确率、正确试次平均反应时与容量 K；
//                      末尾再附每个 set size 的汇总行（StudyDuration 记为「汇总」）；
//                      超时试次不纳入 K 计算，数量见 TotalTrials / TimeoutTrials / ValidTrials；
//                      低正确率条件行以 LowAccRed = 1 标记（标红受 XLSX 版本限制）
//   Sheet3「Grouped Trials」：列与 Sheet1 相同，但按「呈现时长 → set size」分组重排，
//                            同一条件的试次在表中连续排列，便于逐条件核对原始数据
//   K = (Acc * N^2 - N) / (N - 1)
//
// 标红能力说明：
//   当前 vendor 内嵌的 SheetJS 0.18.5 不支持单元格样式写入，
//   K 值表用「LowAccRed」列记录判定结果（1 = 应标红），将来升级 XLSX 或
//   改用 ExcelJS 后 markKRowsRed() 会自动生效，调用方无需改动。
//
// 颜色规则（每个试次由 buildProbePlan 一次算好并冻结，drawPlanSquares 只按计划绘制）：
//   记忆阵列：从 9 色调色板中不重复地随机取 setSize 个颜色
//   探测阵列：复制记忆阵列颜色，只替换「变化方块」一个 → 有且只有一个色块变色
//   新颜色  ：只能从「记忆阵列中没出现过的颜色」里挑（严格排除，无兜底放宽）
//   计划生成后立即自检 verifyProbePlan()，违反则 console.error 报警（probePlanViolations）
//   探测阵列直接使用记忆阵列那一份 plan，不会因重入/并发读到别的试次的数据
//
// 鼠标反应：只用鼠标左键直接点击探测阵列中的色块
//   只有左键点击落在色块范围内才算作反应。
//   实验画布仍禁用右键菜单，仅用于防止菜单遮挡刺激，与作答无关。
//
// 变化颜色规则：
//   被选方块变为一个未在阵列中出现过的新颜色。
// ============================================================

// 实验名称：仅用于导出 Excel 的文件名前缀（exp1_clt_被试编号.xlsx）
// 改名只需改这一处，文件名与结束页的手动下载链接会同步生效
const EXPERIMENT_NAME = 'exp1_clt_compare';
const EXPERIMENT_VERSION = 'clt_compare';
const SESSION_ID = (globalThis.crypto && typeof globalThis.crypto.randomUUID === 'function')
    ? globalThis.crypto.randomUUID()
    : 'session-' + Date.now() + '-' + Math.random().toString(16).slice(2);

// 运行版本标记（与 exp1_clt.html 中 exp1_clt.js?v= 保持一致）：
// 用于确认浏览器实际加载的是最新代码，而非缓存里的旧版本
const EXP1_CLT_BUILD = '20261125compare-clean2';
console.log('[exp1_clt] build ' + EXP1_CLT_BUILD + ' loaded');

// 页面版本核对：与 html 的 PAGE_BUILD 不一致时告警
console.log('[exp1_clt] build ' + EXP1_CLT_BUILD + ' loaded');

// 版本自检：页面（exp1_clt.html）与脚本（exp1_clt_compare.js）版本号必须一致。
// 不一致说明浏览器缓存了旧 exp1_clt_compare.js（实验会跑出旧规则、旧导出文件名），
// 在设置页直接给出告警，避免又用旧代码跑一整轮实验。
(function checkBuildMatch() {
    const pageBuild = window.EXP1_CLT_PAGE_BUILD;
    if (pageBuild && pageBuild !== EXP1_CLT_BUILD) {
        console.warn('[exp1_clt] 版本不一致：页面 ' + pageBuild + ' / 脚本 ' + EXP1_CLT_BUILD +
            ' —— 浏览器可能缓存了旧 exp1_clt.js，请 Ctrl+F5 强制刷新');
        const el = document.getElementById('build-mismatch');
        if (el) el.classList.remove('hidden');
    }
})();

const storage = window.CLTStorage;
let p = {};
let stimData = [];
let practiceData = [];
let currentBlock = 1;
let currentTrial = 0;
let blockTrials = [];
let experimentPhase = 'setup';

// --- 断点续做：PRNG 基准与续做状态（2026-10-10，与 exp1_fcdt 结构一致）---
// buildFormalTrials 每次都先把 PRNG 重置回正式试次基准起点，
// 这样无论是否跳过练习，80 个正式试次的顺序都完全一致。
let formalPrngBaseSeed = 0;
let isResumedSession = false;
let resumedFromBlock = 0;   // 0 = 非续做；N = 从第 N 组（block）续做
let resumeCount = 0;        // 本被试累计续做次数（第 1 次续做 = 1）
// 续做起点：若检测到上次未完成的进度，从第 N 组（block）继续；首次（或重新开始）为 1。
let resumeStartBlock = 1;
let currentNItems = 4; // 当前试次的 set size，用于鼠标命中范围判断
let lowAccBlockCount = 0;        // 正确率低于 50% 的 block 累计数（只统计，不终止实验）
// 低正确率标记：set size = 4 的平均正确率低于 50% 时为 true，导出文件名最前面加 F。
// 注意：实验本身不再因低正确率提前结束，永远跑完 9 个 block。
let ss4LowAccuracyFlag = false;  // 实际语义：低正确率标记（SS3/SS4 在 ≥200ms 正确率 < 50%）
let lowAccDetail = { triggered: false, details: [] };  // 各 set size 的判定明细，供 K 值表标红使用

// --- 实验参数（对应 CLT.m getPreferences） ---
const prefs = {
    setSizes: [4, 6],
    studyDurations: [250],                 // compare 版本仅一个呈现时长条件
    numBlocks: 4,
    numTrials: 20,
    nEachCondition: 40,      // 每个 set size 40 trial，共 80 trial
    nPractice: 8,           // 调试练习 8 trial，2 种 set size 各 4 trial
    ITI: 500,               // ms
    retentionInterval: 1000,// ms
    stimSize: 51,
    minDist: 51 * 2.5,
    fixationSize: 8,
    breakLength: 0.25,      // 分钟 = 15 秒（block 间强制休息，与 exp1_fcdt 保持一致）
    labelFontSize: 18
};

// 调试模式只缩短实际运行的正式实验；prefs 始终保留完整科研设计参数，便于自检。
let runtimeNumBlocks = prefs.numBlocks;
let runtimeNumTrials = prefs.numTrials;
let runtimeBreakLength = prefs.breakLength;

// --- 兼容旧版数据结构的内部标签映射 ---
// 鼠标版不再绘制数字标签，但保留映射字段，便于旧表结构继续接收 response/correctResponse。
const LABEL_CHARS = ['1', '2', '3', '8', '9', '0'];
let trialLabels = [];

const colors_9 = [
    [255, 0, 0], [0, 255, 0], [0, 0, 255], [255, 255, 0],
    [255, 0, 255], [0, 255, 255], [255, 255, 255], [0, 0, 0], [255, 128, 0]
];

// --- PRNG ---
let seed = 1;
function seedRandom(s) { seed = s; }
function random() {
    let x = Math.sin(seed++) * 10000;
    return x - Math.floor(x);
}
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}
function randInt(min, max) {
    return Math.floor(random() * (max - min + 1)) + min;
}

// --- 高精度计时器 ---
function wait(ms) {
    return new Promise(resolve => {
        const start = performance.now();
        function check(time) {
            if (time - start >= ms) { resolve(); }
            else { requestAnimationFrame(check); }
        }
        requestAnimationFrame(check);
    });
}

// --- DOM 元素 ---
const screens = {
    setup: document.getElementById('setup-phase'),
    subjectid_confirm: document.getElementById('subjectid-confirm-phase'),
    instructions: document.getElementById('instructions-phase'),
    experiment: document.getElementById('experiment-phase'),
    break: document.getElementById('break-phase'),
    upload: document.getElementById('upload-phase'),
    finished: document.getElementById('finished-phase')
};
const canvas = document.getElementById('expCanvas');
const ctx = canvas.getContext('2d', { alpha: false });

// --- SubjectID 自动生成 ---
// 规则：{subGroup}_{3位点击顺序}_{3位随机后缀}
// 调试模式：subGroup=0 时固定为 1111
const subjectIDInput = document.getElementById('subjectID');
let generatedClickOrder = null;

function generateSubjectID(subGroupValue) {
    if (String(subGroupValue) === '0') return '1111';
    const order = String(Math.floor(Math.random() * 1000)).padStart(3, '0');
    const alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789';
    const suffix = Array.from({ length: 3 }, () => alphabet[Math.floor(Math.random() * alphabet.length)]).join('');
    generatedClickOrder = order;
    return `${subGroupValue}_${order}_${suffix}`;
}

function formatDateYYYYMMDD(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}${m}${day}`;
}

// 掩蔽身份证号（用于数据导出，保护隐私）—— 本版本不再采集，保留 helper 以防旧引用
function maskIdCard(idCard) {
    if (!idCard) return '';
    if (idCard.length > 10) {
        return idCard.substring(0, 4)
            + '*'.repeat(idCard.length - 10)
            + idCard.substring(idCard.length - 6);
    }
    return idCard;
}

// 全角数字 → 半角数字
function normalizeDigits(v) {
    return String(v == null ? '' : v)
        .trim()
        .replace(/[\uFF10-\uFF19]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFF10 + 0x30));
}

// 组别：
//   0 = 调试模式
//   1 = 先做变化觉察任务（FCDT）再做变化定位任务（CLT）
//   2 = 先做变化定位任务（CLT）再做变化觉察任务（FCDT）
// 两个实验程序共用同一套编号，导出列名为 subGroup / Order，合并两个范式的数据后
// 即可检验「Order × ExperimentType」的交互（即顺序效应）。
const GROUP_LABELS = {
    0: '0 = 调试模式',
    1: '1 = 先做变化觉察任务（FCDT），再做变化定位任务（CLT）',
    2: '2 = 先做变化定位任务（CLT），再做变化觉察任务（FCDT）'
};

function readSubGroup() {
    const input = document.getElementById('subGroup');
    return normalizeDigits(input ? input.value : '');
}

// --- 设置表单 ---
document.getElementById('start-btn').addEventListener('click', async () => {
    const form = document.getElementById('setup-form');
    if (!form.reportValidity()) return;

    const subGroupValue = readSubGroup();
    if (!/^[012]$/.test(subGroupValue)) {
        alert('组别填写有误，请按主试安排填写：\n' + GROUP_LABELS[0] + '\n' + GROUP_LABELS[1] + '\n' + GROUP_LABELS[2]);
        const subGroupInput = document.getElementById('subGroup');
        if (subGroupInput && subGroupInput.focus) subGroupInput.focus();
        return;
    }
    p.subGroup = Number(subGroupValue);

    const subjectID = generateSubjectID(p.subGroup);
    subjectIDInput.value = subjectID;
    p.subjectID = subjectID;

    // 调试模式：subGroup = 0
    p.debugMode = p.subGroup === 0;
    if (p.debugMode) {
        console.log('[exp1_clt] 调试模式已触发（组别 = 0）');
    } else {
        console.log('[exp1_clt] 正式模式（组别=' + p.subGroup + '，被试编号=' + subjectID + '）');
    }

    p.saveLocal = true;
    p.expDate = formatDateYYYYMMDD(new Date());

    // 调试模式固定种子，正式模式用被试编号派生
    p.rndSeed = p.debugMode ? 1111 : Number(subjectID) || Date.now();
    seedRandom(p.rndSeed);

    runtimeNumBlocks = prefs.numBlocks;
    runtimeNumTrials = prefs.numTrials;
    runtimeBreakLength = prefs.breakLength;
    if (p.debugMode) {
        runtimeNumBlocks = 1;
        runtimeNumTrials = 5;
        runtimeBreakLength = 0;
        p.subjectID = '1111';
        p.rndSeed = 1111;
        subjectIDInput.value = '1111';
    }

    // 正式试次序列使用专属 PRNG 种子，与练习阶段完全解耦
    formalPrngBaseSeed = p.rndSeed + 7919;

    if (!canvas.requestPointerLock) {
        showSetupError('当前浏览器不支持鼠标指针锁定，请使用最新版 Chrome 或 Edge。');
        return;
    }
    let lockPromise;
    try {
        lockPromise = canvas.requestPointerLock();
    } catch (err) {
        showSetupError('无法锁定鼠标指针，请允许浏览器锁定鼠标后重试。');
        return;
    }
    let fullscreenPromise = null;
    try {
        if (document.documentElement.requestFullscreen) {
            fullscreenPromise = document.documentElement.requestFullscreen();
        }
    } catch (err) { console.warn('Fullscreen denied.', err); }
    if (lockPromise && typeof lockPromise.then === 'function') {
        try { await lockPromise; } catch (err) {
            showSetupError('浏览器未允许锁定鼠标指针，实验未开始。请重新点击开始并允许锁定。');
            return;
        }
    }
    if (document.pointerLockElement !== canvas) {
        const locked = await new Promise(resolve => {
            let settled = false;
            const finish = value => {
                if (settled) return;
                settled = true;
                document.removeEventListener('pointerlockchange', onChange);
                clearTimeout(timer);
                resolve(value);
            };
            const onChange = () => finish(document.pointerLockElement === canvas);
            const timer = setTimeout(() => finish(false), 1500);
            document.addEventListener('pointerlockchange', onChange);
        });
        if (!locked) {
            showSetupError('鼠标指针锁定未生效，实验未开始。请重新点击开始并允许锁定。');
            return;
        }
    }
    if (fullscreenPromise && typeof fullscreenPromise.then === 'function') {
        try { await fullscreenPromise; } catch (err) { console.warn('Fullscreen denied.', err); }
    }

    try {
        if (!storage) throw new Error('Storage adapter is unavailable');
        await storage.initializeStorage({
            participant: { ...p }, build: EXP1_CLT_BUILD, experiment: EXPERIMENT_NAME
        });
    } catch (err) {
        console.error('[CLT storage] initialization failed:', err);
        if (storage && storage.backendName === 'jatos') {
            showSetupError('JATOS 初始化失败，实验尚未开始。请联系实验管理员。');
            return;
        }
    }

    let resumable = null;
    try {
        resumable = await storage.loadProgress();
    } catch (err) {
        console.warn('[' + EXPERIMENT_NAME + '] 读取续做进度失败，将从头开始：', err);
    }
    if (resumable && resumable.completedBlocks > 0 && resumable.completedBlocks < runtimeNumBlocks) {
        resumeStartBlock = resumable.completedBlocks + 1;
        isResumedSession = true;
        resumedFromBlock = resumeStartBlock;
        resumeCount = (resumable.resumeCount || 0) + 1;
        console.log('[' + EXPERIMENT_NAME + '] 检测到未完成的进度（已完成 ' +
            resumable.completedBlocks + ' 组），将从第 ' + resumeStartBlock + ' 组续做。');
    }

    // 被试编号确认页：让被试看清并记录编号
    const confirmIdEl = document.getElementById('confirm-subject-id');
    if (confirmIdEl) confirmIdEl.innerText = '被试编号：' + p.subjectID;
    switchScreen('subjectid-confirm');
    experimentPhase = 'subjectid_confirm';
});

function showSetupError(message) {
    const el = document.getElementById('build-mismatch');
    if (!el) return;
    el.innerText = message;
    el.classList.remove('hidden');
}

function switchScreen(screenName) {
    for (let key in screens) screens[key].classList.add('hidden');
    screens[screenName].classList.remove('hidden');
}

// --- 全局键盘监听：仅用于阶段切换，不再用于试次作答 ---
window.addEventListener('keydown', (e) => {
    if (e.repeat) return;

    if (experimentPhase === 'subjectid_confirm' && (e.code === 'Space' || e.key === ' ')) {
        e.preventDefault();
        switchScreen('instructions');
        experimentPhase = 'instructions';
        return;
    }
    if (experimentPhase === 'instructions' && (e.code === 'Space' || e.key === ' ')) {
        e.preventDefault();
        if (resumeStartBlock > 1) startFormal();
        else startPractice();
    } else if (experimentPhase === 'practice_end' && (e.code === 'Space' || e.key === ' ')) {
        e.preventDefault();
        startFormal();
    } else if (experimentPhase === 'break_countdown' || experimentPhase === 'break_wait') {
        if (experimentPhase === 'break_wait' && (e.code === 'Space' || e.key === ' ')) {
            e.preventDefault();
            nextBlock();
        }
    }
});

// ============================================================
//  防误触层（guard rails）
//  目的：避免被试或浏览器在实验进行中触发以下操作导致实验中断 / 数据丢失：
//   1) 浏览器刷新、关闭、离开（beforeunload）
//   2) 浏览器后退 / 前进（Backspace、Alt+← / Alt+→、popstate）
//   3) 快捷键被劫持（F5 / Ctrl+R / Ctrl+W / Ctrl+T / F12 / DevTools）
//   4) Esc 退出全屏（屏蔽 Esc）
//   5) 鼠标长按 → 系统右键菜单 / 拖拽（mousedown 持续触发、dragstart、selectstart）
//   6) 触屏 / 滚轮 / 缩放（touchstart、wheel、gesturestart）
//   7) 页面失焦 / 切到后台（visibilitychange 给出告警但不清空数据）
//   8) 页面被 resize（仅在 #experiment-phase 锁住尺寸，避免破坏注视点位置）
//  实验进行中：仅按每个阶段规定的"有效键"（SPACE / 鼠标左键）才生效；
//  其他任何键 / 鼠标 / 触屏 / 系统事件一律被静默 preventDefault，实验流程不受影响。
// ============================================================

// 每个阶段被允许的键盘按键集合（其他键一律屏蔽）
const ALLOWED_KEYS = new Set([' ', 'Space', 'Spacebar']);

// 实验中需要拦截的系统级组合键（不拦截时浏览器会触发默认行为）
const BLOCKED_SHORTCUTS = new Set([
    'F5', 'F12',                                                  // 刷新 / 开发者工具
    'F1', 'F2', 'F3', 'F4', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11' // 其他功能键
]);
function isBlockedShortcut(e) {
    if (BLOCKED_SHORTCUTS.has(e.key)) return true;
    if (e.ctrlKey || e.metaKey) {
        // 拦截 Ctrl/Cmd + R / W / T / N / P / F / S / Shift+Delete 等
        if (/^[rwnpfst]$/i.test(e.key)) return true;
        if (e.key === 'F4' && e.altKey) return true;   // Alt+F4 关窗
    }
    if (e.altKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight' || e.key === 'Backspace')) {
        return true;                                       // Alt+←/→ 浏览器后退 / 前进
    }
    return false;
}

// (1) 全局 keydown 第二层：仅放行 ALLOWED_KEYS，其余静默屏蔽
window.addEventListener('keydown', (e) => {
    if (e.repeat) return;                                 // 防止长按重复触发

    // 关键例外：焦点落在输入框 / 下拉框里时（setup 阶段填姓名 / 性别 / 年龄 /
    // 身份证 / 手机号），防误触层必须放行普通按键——在 capture 阶段对 keydown
    // 调用 preventDefault() 会阻止字符被插入输入框，被试将一个字都打不进去。
    const inFormField = (e.target instanceof HTMLInputElement)
        || (e.target instanceof HTMLTextAreaElement)
        || (e.target instanceof HTMLSelectElement);

    if (isBlockedShortcut(e)) {
        e.preventDefault();
        e.stopPropagation();
        return;
    }
    if (inFormField) return;                               // 表单输入：放行所有普通按键
    if (e.key === 'Escape') {                              // 屏蔽 Esc 退出全屏
        e.preventDefault();
        return;
    }
    if (e.key === 'Backspace' && !(e.target instanceof HTMLInputElement)) {
        // Backspace 在非输入框下被浏览器解读为"后退" → 屏蔽
        e.preventDefault();
        return;
    }
    // 其他所有非放行键：静默屏蔽
    if (!ALLOWED_KEYS.has(e.key) && !ALLOWED_KEYS.has(e.code) && e.key !== 'Tab') {
        e.preventDefault();
    }
}, true);                                                  // capture 阶段优先

// (2) 屏蔽所有元素的右键菜单（不仅 canvas）
//     实验阶段 canvas 上的右键必须屏蔽（菜单会遮挡刺激、干扰作答）；
//     但 setup 阶段表单里的右键菜单保留（便于粘贴身份证号 / 手机号），
//     文本输入框的右键菜单里没有任何导航项，不会导致实验中断。
window.addEventListener('contextmenu', (e) => {
    const t = e.target;
    if (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement) return;
    e.preventDefault();
});

// (3) 屏蔽拖拽 / 选中 / 复制（避免误选文字拖出浏览器）
window.addEventListener('dragstart', (e) => e.preventDefault());
window.addEventListener('selectstart', (e) => {
    // 允许在 setup 阶段（输入框）选中文字；其他阶段屏蔽
    const phase = experimentPhase;
    if (phase !== 'setup' && phase !== 'instructions') {
        if (!(e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement)) {
            e.preventDefault();
        }
    }
});
window.addEventListener('copy', (e) => {
    if (!(e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement)) {
        e.preventDefault();
    }
});

// (4) 屏蔽滚轮 / 触屏手势（避免页面缩放 / 滚动 / 系统菜单）
window.addEventListener('wheel', (e) => {
    if (e.ctrlKey) e.preventDefault();                    // Ctrl+滚轮 = 缩放
}, { passive: false });
window.addEventListener('touchstart', (e) => {
    if (e.touches.length > 1) e.preventDefault();         // 多指 = 缩放手势
}, { passive: false });
window.addEventListener('gesturestart', (e) => e.preventDefault());

// (5) 屏蔽全屏切换 / 重新加载（beforeunload）：实验进行中给出离开提示
window.addEventListener('beforeunload', (e) => {
    const phase = experimentPhase;
    if (phase === 'setup' || phase === 'finished') return;  // 未开始 / 已结束：允许直接离开
    e.preventDefault();
    e.returnValue = '实验正在进行中，关闭或刷新页面将导致数据丢失。确定要离开吗？';
    return e.returnValue;
});

// (6) 页面失焦 / 切到后台：给出告警，但**不**清空数据、不终止试次
let visibilityWarned = false;
document.addEventListener('visibilitychange', () => {
    if (document.hidden && !visibilityWarned) {
        visibilityWarned = true;
        console.warn('[' + EXPERIMENT_NAME + '] 页面失焦：被试切到其他窗口 / 标签页。');
    } else if (!document.hidden) {
        visibilityWarned = false;
    }
});
window.addEventListener('blur', () => {
    if (experimentPhase === 'trial_response' || experimentPhase === 'practice_response') {
        console.warn('[' + EXPERIMENT_NAME + '] 实验进行中窗口失焦，请保持页面在前台。');
    }
});

// (7) popstate（浏览器后退 / 前进）：压入 dummy 历史项并屏蔽后退
history.pushState({ phase: 'experiment' }, '');
window.addEventListener('popstate', () => {
    const phase = experimentPhase;
    if (phase !== 'setup' && phase !== 'finished') {
        // 实验进行中：阻止后退，立即推回一项
        history.pushState({ phase: 'experiment' }, '');
        console.warn('[' + EXPERIMENT_NAME + '] 浏览器后退被屏蔽（实验进行中）。');
    }
});

// (8) 屏蔽浏览器对鼠标的默认处理（防止长按出系统菜单、拖拽光标、指针变化）。
//     mousedown 一次按下只触发一次，浏览器在按住期间不会自动 repeat mousedown；
//     右键菜单已由 (2) contextmenu preventDefault 屏蔽；
//     这里额外阻止 mousedown 的默认光标变化 / 文本选中，并屏蔽 canvas dragstart。
canvas.addEventListener('mousedown', (e) => {
    // 在所有阶段都阻止默认行为（除 setup 阶段输入框，但 canvas 不在输入框里）
    e.preventDefault();
}, true);
canvas.addEventListener('dragstart', (e) => e.preventDefault());
canvas.addEventListener('drag', (e) => e.preventDefault());

// 与作答无关，仅为防止浏览器右键菜单遮挡刺激。
// （已由上面的 (2) 全局 contextmenu 屏蔽；这里保留 canvas 上的显式监听便于阅读）

// 把鼠标坐标换算到 canvas 坐标，并判断点击落在哪个色块上。
function clickedItemIndex(x, y) {
    if (!probeSnapshot || !Array.isArray(probeSnapshot.locs)) return -1;
    const half = prefs.stimSize / 2;
    for (let i = 0; i < currentNItems; i++) {
        const loc = probeSnapshot.locs[i];
        if (Math.abs(x - loc.x) <= half && Math.abs(y - loc.y) <= half) return i;
    }
    return -1;
}

canvas.addEventListener('mousedown', (e) => {
    const phase = experimentPhase;
    if (phase !== 'practice_response' && phase !== 'trial_response') return;
    if (e.button !== 0) return;
    e.preventDefault();

    const itemIdx = clickedItemIndex(virtualCursor.x, virtualCursor.y);
    if (itemIdx < 0) return;
    clickedItemIndexCircle = itemIdx;

    // response/correctResponse 沿用旧版标签槽编号，确保既有数据结构兼容；
    // 实际选择的是 itemIdx 对应的色块，导出时再转换为 1 起始的色块序号。
    const mappedSlot = labelMap.indexOf(itemIdx);
    const responseNum = mappedSlot >= 0 ? mappedSlot + 1 : itemIdx + 1;

    if (phase === 'practice_response') recordPracticeResponse(responseNum);
    else recordResponse(responseNum);
});

// ============================================================
//  练习阶段（15 trial，3 种 set size × 5，呈现时长逐试次随机，全部 change）
// ============================================================
let practiceTrials = [];
let practiceTrialIdx = 0;
let practiceAcc = [];

function startPractice() {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;

    const nPerSzPractice = prefs.nPractice / prefs.setSizes.length;
    practiceTrials = [];
    for (let ss of prefs.setSizes) {
        for (let i = 0; i < nPerSzPractice; i++) {
            const durationIdx = randInt(0, prefs.studyDurations.length - 1);
            practiceTrials.push({
                setSize: ss,
                studyDuration: prefs.studyDurations[durationIdx],
                change: 1
            }); // CLT 全部 change
        }
    }
    shuffleArray(practiceTrials);
    practiceTrialIdx = 0;
    practiceAcc = [];
    practiceData = [];

    switchScreen('experiment');
    runPracticeTrial();
}

let activeColors = [];
let activeLocs = [];
let trialData = {};
let practiceTrialData = {};
let responseStartTime = 0;
let responseTimeoutTimer = null;
const RESPONSE_LIMIT_MS = 4000;

// --- 自绘虚拟光标状态（与 circle 版一致）---
let virtualCursor = { x: 0, y: 0 };
let cursorActive = false;
let cursorSkipNextMove = false;
let cursorStartX = 0, cursorStartY = 0;
let cursorTravelPx = 0;
let lastClientX = null, lastClientY = null;
let pointerLocked = false;
let probeSnapshot = null;
let clickedItemIndexCircle = -1;

// CLT 专用：变化方块索引、新颜色、兼容旧数据的内部标签映射
// 这些全局量是「当前试次」的镜像，记录数据与鼠标判断用；
// 绘制一律走 trialPlan（见 buildProbePlan），不直接读这些全局量。
let changeItemIdx = 0;
let changeColorIdx = 0;
let labelMap = [];
let correctLabel = 0;

// --- 当前试次的「探测计划」---
// 记忆阵列颜色、探测阵列颜色、位置、标签全部由它一次算好并冻结。
// 这样探测阵列必定是「它自己那次记忆阵列」的唯一一个方块变色，
// 不可能出现两个（或更多）方块同时变色。
let trialPlan = null;
let probePlanViolations = 0;   // 自检失败计数，正常情况下恒为 0

// 自检：探测阵列相对记忆阵列必须有且只有一个方块变色，
// 且变化后的颜色不能是记忆阵列中出现过的颜色
function verifyProbePlan(plan) {
    let diff = [];
    for (let i = 0; i < plan.setSize; i++) {
        if (plan.probeColors[i] !== plan.memColors[i]) diff.push(i);
    }
    const onlyOneChanged = diff.length === 1 && diff[0] === plan.changeIdx;
    const newColorNotInMemory = !plan.memColors.includes(plan.probeColors[plan.changeIdx]);
    const colorReallyChanged = plan.probeColors[plan.changeIdx] !== plan.memColors[plan.changeIdx];
    const noDupInMemory = new Set(plan.memColors).size === plan.setSize;
    const ok = onlyOneChanged && newColorNotInMemory && colorReallyChanged && noDupInMemory;
    if (!ok) {
        probePlanViolations++;
        console.error('[exp1_clt] 探测阵列自检未通过（不应发生）：', {
            diff: diff, memColors: plan.memColors, probeColors: plan.probeColors,
            changeIdx: plan.changeIdx, setSize: plan.setSize, violations: probePlanViolations
        });
    }
    return ok;
}

// 生成一次试次的完整视觉计划
//   记忆阵列：从 9 色调色板中不重复地随机取 setSize 个颜色
//   探测阵列：复制记忆阵列颜色，只把「变化方块」换成新颜色 → 恰好 1 个方块变色
//   新颜色  ：只能从「记忆阵列中没出现过的颜色」里挑，严格排除，无兜底放宽
function buildProbePlan(setSize) {
    let colorIndexArr = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    shuffleArray(colorIndexArr);
    const memColors = colorIndexArr.slice(0, setSize);

    const locs = getStimLocs(setSize, canvas.width / 2, canvas.height / 2);

    // 变化方块：随机 1 个
    const changeIdx = randInt(0, setSize - 1);

    // 新颜色：排除记忆阵列出现过的全部颜色
    let newColorPool = [];
    for (let c = 0; c < colors_9.length; c++) {
        if (!memColors.includes(c)) newColorPool.push(c);
    }
    if (newColorPool.length === 0) {
        // 调色板共 9 色、setSize 最大 6，实际不可能触发；
        // 保留防御但不放宽条件（绝不退回记忆阵列已用的颜色）
        throw new Error('buildProbePlan: 调色板颜色不足，无法选出不属于记忆阵列的新颜色');
    }
    const newColor = newColorPool[randInt(0, newColorPool.length - 1)];

    // 探测阵列颜色：仅替换变化方块那一项
    const probeColors = memColors.slice();
    probeColors[changeIdx] = newColor;

    // 内部标签映射：鼠标版不绘制标签，仅用于保持旧数据字段的编号语义。
    let lmap = [];
    for (let i = 0; i < setSize; i++) lmap.push(i);
    shuffleArray(lmap);

    // 兼容旧数据的标签字符元数据（不显示给被试）
    let charPool = [0, 1, 2, 3, 4, 5];
    shuffleArray(charPool);

    const plan = {
        setSize: setSize,
        locs: locs,
        memColors: memColors,           // 记忆阵列颜色
        probeColors: probeColors,       // 探测阵列颜色（恰好 1 项与 memColors 不同）
        changeIdx: changeIdx,           // 变化方块的索引
        origColor: memColors[changeIdx],// 变化方块原色
        newColor: newColor,             // 变化方块新色（不在 memColors 中）
        labelMap: lmap,
        trialLabels: charPool.slice(0, setSize),
        correctLabel: lmap.indexOf(changeIdx) + 1
    };

    verifyProbePlan(plan);
    return plan;
}

// 把试次计划同步到全局镜像（记录数据 / 按键判断用；绘制不依赖这些全局量）
function applyPlanToGlobals(plan) {
    activeColors = plan.memColors.slice();
    activeLocs = plan.locs;
    changeItemIdx = plan.changeIdx;
    changeColorIdx = plan.newColor;
    labelMap = plan.labelMap.slice();
    trialLabels = plan.trialLabels.slice();
    correctLabel = plan.correctLabel;
    currentNItems = plan.setSize;
}

// 按指定颜色数组绘制当前计划的全部方块（记忆阵列 / 探测阵列共用）
function drawPlanSquares(plan, colorIdxArr) {
    for (let i = 0; i < plan.setSize; i++) {
        drawSquare(plan.locs[i].x, plan.locs[i].y, colors_9[colorIdxArr[i]]);
    }
}

async function runPracticeTrial() {
    if (practiceTrialIdx >= practiceTrials.length) {
        endPractice();
        return;
    }

    experimentPhase = 'practice_running';
    let condition = practiceTrials[practiceTrialIdx];
    currentNItems = condition.setSize;
    practiceTrialData = {
        block: 0,
        trial: practiceTrialIdx + 1,
        setSize: condition.setSize,
        studyDuration: condition.studyDuration,
        isChange: 1
    };

    // ITI
    drawBackground(); drawFixation();
    await wait(prefs.ITI);

    await showStudyArrayPractice(condition);
}

async function showStudyArrayPractice(condition) {
    // 一次算好并冻结本试次的视觉计划：记忆阵列 / 探测阵列 / 位置 / 标签
    // 计划内已保证「有且只有一个方块变色」且「新颜色不在记忆阵列中出现」
    const plan = buildProbePlan(condition.setSize);
    trialPlan = plan;
    applyPlanToGlobals(plan);

    // 学习阵列（仅方块，无标签）
    drawBackground(); drawFixation();
    drawPlanSquares(plan, plan.memColors);
    const studyStartTime = performance.now();
    await wait(condition.studyDuration);
    practiceTrialData.studyDurationActual = (performance.now() - studyStartTime).toFixed(2);

    // 保留间隔（注视点全程固定显示，不消失）
    drawBackground(); drawFixation();
    await wait(prefs.retentionInterval);

    // 探测阵列：直接用同一个 plan 的 probeColors，任何情况下都只有一个方块变色
    showTestArrayPractice(plan);
}

function showTestArrayPractice(plan) {
    probeSnapshot = { locs: plan.locs, colorIdx: plan.probeColors };
    clickedItemIndexCircle = -1;
    resetVirtualCursor();
    redrawProbeScreen();

    experimentPhase = 'practice_response';
    responseStartTime = performance.now();
    clearTimeout(responseTimeoutTimer);
    responseTimeoutTimer = setTimeout(() => {
        if (experimentPhase === 'practice_response') recordPracticeResponse(null);
    }, RESPONSE_LIMIT_MS);
}

function recordPracticeResponse(responseNum) {
    // 立即锁定反应阶段：防止鼠标连击在同一试次内触发第二次响应
    if (experimentPhase !== 'practice_response') return;
    experimentPhase = 'practice_locked';
    clearTimeout(responseTimeoutTimer);
    responseTimeoutTimer = null;
    hideVirtualCursor();

    const timedOut = responseNum === null;
    let rt = timedOut ? null : (performance.now() - responseStartTime) / 1000;
    practiceTrialData.rt = rt;
    practiceTrialData.response = responseNum;          // 超时为空值 null，否则为标签槽编号（1..nItems）
    practiceTrialData.timeout = timedOut;
    practiceTrialData.correctResponse = correctLabel;
    practiceTrialData.changeItem = changeItemIdx;
    practiceTrialData.origColorIdx = activeColors[changeItemIdx];
    practiceTrialData.probeColorIdx = changeColorIdx;
    practiceTrialData.labelMap = JSON.stringify(labelMap);
    practiceTrialData.trialLabels = JSON.stringify(trialLabels.map(i => LABEL_CHARS[i]));
    practiceTrialData.labelChars = LABEL_CHARS.join('');
    practiceTrialData.accuracy = !timedOut && responseNum === correctLabel ? 1 : 0;

    practiceAcc.push(practiceTrialData.accuracy);
    practiceData.push(Object.assign({}, practiceTrialData));

    drawBackground(); drawFixation();
    practiceTrialIdx++;
    setTimeout(() => runPracticeTrial(), 50);
}

function endPractice() {
    let accPct = Math.round(practiceAcc.reduce((a, b) => a + b, 0) / practiceAcc.length * 100);
    experimentPhase = 'practice_end';
    drawBackground();

    ctx.fillStyle = 'rgb(255,255,255)';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = 'bold 36px Inter, sans-serif';
    ctx.fillText('练习结束', canvas.width / 2, canvas.height / 2 - 80);
    ctx.font = '28px Inter, sans-serif';
    ctx.fillText(`正确率：${accPct}%`, canvas.width / 2, canvas.height / 2 - 20);
    ctx.font = '20px Inter, sans-serif';
    ctx.fillStyle = 'rgb(255,235,59)';
    ctx.fillText('按 空格键(Space) 开始正式实验', canvas.width / 2, canvas.height / 2 + 60);
}

// ============================================================
//  正式实验阶段
// ============================================================
// 生成完整正式试次序列：set size 4/6（固定 250 ms）各 40 次，整体只洗牌一次。
// 关键不变量（与 exp1_fcdt 一致）：函数开头把 PRNG 重置回 formalPrngBaseSeed，
// 因此"跳过练习直接续做"与"正常做完练习再进正式"得到的 80 个正式试次序列完全相同。
// 若正式基准起点尚未捕获（formalPrngBaseSeed 为 0），退回用 p.rndSeed 作为起点；
// p.rndSeed 未设置时（如测试沙箱直接调用 buildFormalTrials）再用 Date.now() 兜底，
// 避免 seedRandom(undefined) 污染 PRNG 状态导致 shuffle 无位移。
function buildFormalTrials() {
    const baseSeed = formalPrngBaseSeed !== 0 ? formalPrngBaseSeed : (p.rndSeed || Date.now());
    seedRandom(baseSeed);
    const trials = [];
    for (let ss of prefs.setSizes) {
        for (let duration of prefs.studyDurations) {
            for (let i = 0; i < prefs.nEachCondition; i++) {
                trials.push({ setSize: ss, studyDuration: duration, change: 1 });
            }
        }
    }
    shuffleArray(trials);
    return trials;
}

// 生成后自检：总数、各条件计数和按 50 个切分的 block 结构都必须符合设计。
function verifyFormalTrials(trials) {
    const expectedTotal = prefs.setSizes.length * prefs.studyDurations.length * prefs.nEachCondition;
    const conditionCounts = new Map();
    for (let trial of trials) {
        const key = trial.setSize + '|' + trial.studyDuration;
        conditionCounts.set(key, (conditionCounts.get(key) || 0) + 1);
    }

    let conditionsValid = true;
    for (let ss of prefs.setSizes) {
        for (let duration of prefs.studyDurations) {
            const key = ss + '|' + duration;
            if (conditionCounts.get(key) !== prefs.nEachCondition) conditionsValid = false;
        }
    }

    const blocks = [];
    for (let start = 0; start < trials.length; start += prefs.numTrials) {
        blocks.push(trials.slice(start, start + prefs.numTrials));
    }
    const blocksValid = blocks.length === prefs.numBlocks &&
        blocks.every(block => block.length === prefs.numTrials);
    const valid = trials.length === expectedTotal && conditionsValid && blocksValid;

    if (!valid) {
        console.error('[exp1_clt] 正式试次生成自检未通过（不应发生）：', {
            total: trials.length,
            expectedTotal: expectedTotal,
            conditionCounts: Object.fromEntries(conditionCounts),
            blockSizes: blocks.map(block => block.length),
            expectedBlocks: prefs.numBlocks,
            expectedTrialsPerBlock: prefs.numTrials
        });
    } else {
        console.log('[exp1_clt] 正式试次生成自检通过：' +
            (prefs.setSizes.length * prefs.studyDurations.length) + ' 种条件各 ' + prefs.nEachCondition +
            ' 次，共 ' + trials.length + ' 次，' +
            prefs.numBlocks + ' 个 block × ' + prefs.numTrials + ' 次');
    }
    return valid;
}

function startFormal() {
    allFormalTrials = buildFormalTrials();
    verifyFormalTrials(allFormalTrials);

    // 续做：跳过练习、从上次完成后的下一组开始；否则正常从第 1 组开始。
    currentBlock = Math.min(resumeStartBlock, runtimeNumBlocks);
    if (currentBlock > 1) {
        console.log('[' + EXPERIMENT_NAME + '] 从第 ' + currentBlock + ' 组续做（上次已完成 ' + (currentBlock - 1) + ' 组）');
    }
    startBlock();
}

let allFormalTrials = [];
let formalTrialGlobalIdx = 0;

function startBlock() {
    blockTrials = allFormalTrials.slice(
        (currentBlock - 1) * runtimeNumTrials,
        currentBlock * runtimeNumTrials
    );
    currentTrial = 0;
    formalTrialGlobalIdx = (currentBlock - 1) * runtimeNumTrials;

    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
    switchScreen('experiment');
    runTrial();
}

async function runTrial() {
    if (currentTrial >= blockTrials.length) {
        endBlock();
        return;
    }

    experimentPhase = 'trial_running';
    let condition = blockTrials[currentTrial];
    currentNItems = condition.setSize;
    trialData = {
        subjectID: p.subjectID,
        experimentVersion: EXPERIMENT_VERSION,
        sessionID: SESSION_ID,
        block: currentBlock,
        trial: currentTrial + 1,
        setSize: condition.setSize,
        studyDuration: condition.studyDuration,
        isChange: 1,
        experimentType: EXPERIMENT_NAME
    };

    // ITI
    drawBackground(); drawFixation();
    await wait(prefs.ITI);

    await showStudyArray(condition);
}

async function showStudyArray(condition) {
    // 一次算好并冻结本试次的视觉计划：记忆阵列 / 探测阵列 / 位置 / 标签
    // 计划内已保证「有且只有一个方块变色」且「新颜色不在记忆阵列中出现」
    const plan = buildProbePlan(condition.setSize);
    trialPlan = plan;
    applyPlanToGlobals(plan);

    trialData.itemLocs = JSON.stringify(plan.locs.map(l => [Math.round(l.x), Math.round(l.y)]));
    trialData.itemColors = JSON.stringify(plan.memColors);

    // 学习阵列（仅方块，无标签）
    drawBackground(); drawFixation();
    drawPlanSquares(plan, plan.memColors);
    const studyStartTime = performance.now();
    await wait(condition.studyDuration);
    trialData.studyDurationActual = (performance.now() - studyStartTime).toFixed(2);

    // 保留间隔（注视点全程固定显示，不消失）
    drawBackground(); drawFixation();
    await wait(prefs.retentionInterval);

    trialData.changeItem = plan.changeIdx;
    trialData.origColorIdx = plan.origColor;
    trialData.probeColorIdx = plan.newColor;
    trialData.labelMap = JSON.stringify(plan.labelMap);
    trialData.trialLabels = JSON.stringify(plan.trialLabels.map(i => LABEL_CHARS[i]));
    trialData.labelChars = LABEL_CHARS.join('');
    trialData.correctResponse = plan.correctLabel;

    // 探测阵列：直接用同一个 plan 的 probeColors，任何情况下都只有一个方块变色
    showTestArray(plan);
}

function showTestArray(plan) {
    probeSnapshot = { locs: plan.locs, colorIdx: plan.probeColors };
    clickedItemIndexCircle = -1;
    resetVirtualCursor();
    redrawProbeScreen();

    experimentPhase = 'trial_response';
    responseStartTime = performance.now();
    clearTimeout(responseTimeoutTimer);
    responseTimeoutTimer = setTimeout(() => {
        if (experimentPhase === 'trial_response') recordResponse(null);
    }, RESPONSE_LIMIT_MS);
}

function recordResponse(responseNum) {
    // 立即锁定反应阶段：防止鼠标连击在同一试次内触发第二次响应
    if (experimentPhase !== 'trial_response') return;
    experimentPhase = 'trial_locked';
    clearTimeout(responseTimeoutTimer);
    responseTimeoutTimer = null;
    hideVirtualCursor();

    const timedOut = responseNum === null;
    let rt = timedOut ? null : (performance.now() - responseStartTime) / 1000;
    trialData.rt = rt;
    trialData.response = responseNum;                  // 超时为空值 null，否则为标签槽编号（1..nItems）
    trialData.timeout = timedOut;
    trialData.accuracy = !timedOut && responseNum === correctLabel ? 1 : 0;

    // 记录被试选择颜色与正确颜色的 RGB（色块无固定数字标签，以颜色为准）
    const clickedItemIdx = clickedItemIndexCircle;
    const clickedColorIdx = (clickedItemIdx >= 0 && clickedItemIdx < probeColorArr.length)
        ? probeColorArr[clickedItemIdx]
        : null;
    const correctColorIdx = changeColorIdx;
    trialData.responseColorIdx = clickedColorIdx;
    trialData.correctColorIdx = correctColorIdx;
    trialData.responseRGB = clickedColorIdx !== null ? colors_9[clickedColorIdx].join(',') : null;
    trialData.correctRGB = correctColorIdx !== null ? colors_9[correctColorIdx].join(',') : null;

    // 超时试次标记（append-only 两列，不改动既有列语义）：
    //   TrialStatus = 'valid'   | 'timeout'  —— 人眼在 Trial Data 里可直接按此列筛出超时行
    //   CountedInK  = 1         | 0          —— 0 表示该行不参与 K 值计算（见 buildKRow）
    // 说明：超时行在 Response / RT 两列本来就是空值，但空单元格在 Excel 里容易被当成
    //       "漏填"而与真正的缺失混淆，因此额外写一个显式状态列做区分标记。
    trialData.TrialStatus = timedOut ? 'timeout' : 'valid';
    trialData.CountedInK = timedOut ? 0 : 1;

    const completedTrial = Object.assign({}, trialData);
    stimData.push(completedTrial);
    if (storage) {
        storage.saveTrial(prepareStorageRow(completedTrial)).catch(error => {
            console.error('[CLT storage] trial save failed:', error);
        });
    }

    drawBackground(); drawFixation();
    currentTrial++;
    formalTrialGlobalIdx++;
    setTimeout(() => runTrial(), 50);
}

// ============================================================
//  Block 间休息（强制 ≥15 秒倒计时，倒计时结束后按 SPACE 进入下一组）
// ============================================================
let breakTimer = null;
const BREAK_MIN_SECONDS = 15;   // block 间强制休息 15 秒

// 计算某个 block 的正确率（%）
function blockAccuracyPct(blockNo) {
    let rows = stimData.filter(d => d.block === blockNo);
    if (rows.length === 0) return 0;
    return Math.round(rows.reduce((s, d) => s + d.accuracy, 0) / rows.length * 100);
}

// 低正确率判定（决定文件名是否加 F，也在 K 值表中标红对应行）：
// 对 set size = 3 和 4 分别统计「呈现时长 ≥ 200 ms」试次的正确率，
// 任一个低于 50% 即触发；set size = 6 不参与判定。
// 返回 { triggered, details: [{ setSize, accuracyPct, low, nTotal }] }；
// 数据不足时 triggered 为 false（details 中 nTotal 记 0），调用方据此不触发
const LOW_ACC_MIN_DURATION_MS = 250;   // compare 版本仅有 250 ms 条件
const LOW_ACC_MIN_SET_SIZE = 4;        // 判定只覆盖 set size 3 和 4（不超过 4）
const LOW_ACC_THRESHOLD_PCT = 50;      // 正确率低于 50% 触发

function computeLowAccuracyFlag() {
    const details = [];
    for (const ss of [3, 4]) {
        const rows = stimData.filter(
            d => d.setSize === ss && Number(d.studyDuration) >= LOW_ACC_MIN_DURATION_MS
        );
        const n = rows.length;
        const accPct = n > 0 ? rows.reduce((s, d) => s + d.accuracy, 0) / n * 100 : 0;
        details.push({
            setSize: ss,
            accuracyPct: n > 0 ? accPct : null,
            nTotal: n,
            low: n > 0 ? accPct < LOW_ACC_THRESHOLD_PCT : false
        });
    }
    return { triggered: details.some(d => d.low), details: details };
}

function endBlock() {
    const blockAcc = blockAccuracyPct(currentBlock);

    // 只统计正确率低于 50% 的 block 数量（写入结果表备查）。
    // 已取消「累计 2 个 block 低于 50% 自动结束实验」的规则：
    // 无论正确率高低，实验都完整跑完预定 block，低正确率只在导出文件名上加 F。
    if (blockAcc < 50) lowAccBlockCount++;

    if (storage) {
        storage.saveCheckpoint({
            subjectID: p.subjectID,
            block: currentBlock,
            blockAccuracy: blockAcc,
            completedTrials: stimData.length,
            nextBlock: currentBlock < runtimeNumBlocks ? currentBlock + 1 : null
        }).catch(error => {
            console.error('[CLT storage] checkpoint save failed:', error);
        });

        // 线上暂存：每完成一个 block 把进度写进 JATOS Batch Session（仅 JATOS 后端实现）。
        // 只存"完成到第几组"；身份哈希由存储层自己注入，这里不传任何个人信息。
        if (currentBlock < runtimeNumBlocks) {
            storage.saveProgress({
                completedBlocks: currentBlock,
                resumeCount: isResumedSession ? resumeCount : 0,
                updatedAt: new Date().toISOString()
            });
        }
    }

    if (currentBlock < runtimeNumBlocks) {
        showBreakScreen(blockAcc);
    } else {
        finishExperiment();
    }
}

function showBreakScreen(blockAcc) {
    switchScreen('break');
    experimentPhase = 'break_countdown';
    document.getElementById('rest-acc').innerText = `您的正确率为 ${blockAcc}%`;
    startBreakCountdown();
}

// 强制休息倒计时（结束后显示等待提示，按 SPACE 进入下一组）
function startBreakCountdown() {
    clearInterval(breakTimer);
    breakTimer = null;

    const cdEl = document.getElementById('rest-countdown');
    const waitEl = document.getElementById('rest-wait');
    let timeLeft = Math.max(BREAK_MIN_SECONDS, Math.round(runtimeBreakLength * 60));

    experimentPhase = 'break_countdown';
    cdEl.innerText = timeLeft;
    cdEl.classList.remove('hidden');
    waitEl.classList.add('hidden');

    breakTimer = setInterval(() => {
        timeLeft--;
        if (timeLeft <= 0) {
            clearInterval(breakTimer);
            breakTimer = null;
            cdEl.classList.add('hidden');
            waitEl.classList.remove('hidden');
            experimentPhase = 'break_wait';   // 此时按 SPACE 进入下一组
        } else {
            cdEl.innerText = timeLeft;
        }
    }, 1000);
}

// “继续休息 30 秒”按钮与其 ESC 退出实验功能已删除：
// 倒计时结束后只有「按 SPACE 进入下一组」一种操作。

function nextBlock() {
    clearInterval(breakTimer);
    breakTimer = null;
    currentBlock++;
    startBlock();
}

// 保存前统一补齐被试信息（保证每一行都带完整的被试信息）
function enrichRows(rows) {
    for (let d of rows) {
        Object.assign(d, prepareStorageRow(d));
    }
    return rows;
}

function prepareStorageRow(row) {
    return {
        ...row,
        subGroup: p.subGroup,          // 组别（0/1/2）：0=调试，1/2=正式顺序
        subjectID: p.subjectID,
        rndSeed: p.rndSeed,
        resumedFromBlock: isResumedSession ? resumedFromBlock : 0,
        resumeCount: isResumedSession ? resumeCount : 0
    };
}

// ============================================================
//  数据保存
// ============================================================
async function finishExperiment() {
    switchScreen('upload');
    experimentPhase = 'upload';

    let uploadMsg = "";
    let uploadFailed = false;

    enrichRows(stimData);
    ss4LowAccuracyFlag = computeLowAccuracyFlag();

    if (storage && stimData.length > 0) {
        try {
            const saveResult = await storage.saveFinalResult({
                trials: stimData,
                probePlanViolations: probePlanViolations,
                lowAccuracyFlag: ss4LowAccuracyFlag,
                lowAccuracyBlockCount: lowAccBlockCount
            });
            uploadMsg += saveResult && saveResult.message ? saveResult.message : '数据保存成功。';
            await storage.finishExperiment({ successful: true, message: 'CLT mouse completed' });
            // 实验正常完成：清掉 Batch Session 里的进度，避免下次误判为"上次没做完"
            await storage.clearProgress();
        } catch (error) {
            console.error('[CLT storage] final save failed:', error);
            uploadFailed = true;
            uploadMsg += '服务器上传失败: ' + error.message + '。';
            try {
                await storage.finishExperiment({ successful: false, message: 'CLT mouse persistence failed' });
            } catch (finishError) {
                console.error('[CLT storage] failed to close unsuccessful study:', finishError);
            }
        }
    } else if (!storage) {
        uploadMsg += "本版本未配置服务器上传，仅生成本地 Excel。";
    }

    if (p.saveLocal && stimData.length > 0) {
        try {
            downloadXLSX();
            uploadMsg += "本地 Excel 已生成下载。";
        } catch (e) {
            console.error('[exp1_clt] 本地导出失败:', e);
            uploadFailed = true;
            uploadMsg += "本地 Excel 导出失败：" + e.message + "。";
        }
    }

    document.getElementById('upload-status').innerText = uploadMsg;

    // 结束语画面保持干净：仅在异常时显示状态说明
    const statusEl = document.getElementById('finished-text');
    if (statusEl) {
        if (uploadFailed) {
            statusEl.innerText = uploadMsg;
            statusEl.classList.remove('hidden');
        } else {
            statusEl.innerText = '';
            statusEl.classList.add('hidden');
        }
    }

    setTimeout(() => { switchScreen('finished'); }, 2000);

    try {
        if (document.exitFullscreen && document.fullscreenElement) {
            document.exitFullscreen().catch(() => {});
        }
    } catch (e) {}
}

// ============================================================
//  给 K 值表的「低正确率条件行」标红
//  逐单元格写 style（font color: #C00000 深红），避免整行高亮被
//  部分查看器误判为「整行脏数据」。
//
//  当前 vendor 内嵌的 SheetJS 0.18.5 不支持单元格样式写入，
//  此函数在 0.18.5 下是 no-op（函数开头的 canStyle 检查会直接 return）。
//  K 值表里仍写有「LowAccRed」列记录判定结果（1 = 应标红），
//  将来升级 XLSX 到 0.19+ 或改用 ExcelJS 时本函数会自动生效，
//  调用方（downloadXLSX 内的一行调用）无需改动。
// ============================================================
function markKRowsRed(sheet) {
    // 0.18.5 没有样式能力：不写 cell.s，避免污染数据
    const canStyle = false; // XLSX.utils.sheetStyles 在 0.18.5 下 undefined
    if (!canStyle) return;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    // 找 LowAccRed 列与 K 列的相对索引
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: c })];
        headers[c] = cell && cell.v !== undefined ? String(cell.v) : '';
    }
    const redCol = headers.indexOf('LowAccRed');
    const kCol = headers.indexOf('K_Exclude');
    if (redCol === -1) return;

    for (let r = 1; r <= range.e.r; r++) {
        const flagCell = sheet[XLSX.utils.encode_cell({ r: r, c: redCol })];
        const isLow = flagCell && Number(flagCell.v) === 1;
        if (!isLow) continue;
        // K 值为负的条件行排除（避免把 50ms 拉低正确率导致的负 K 误判为「低正确率」）
        if (kCol !== -1) {
            const kCell = sheet[XLSX.utils.encode_cell({ r: r, c: kCol })];
            if (kCell && Number(kCell.v) < 0) continue;
        }
        // 整行标红
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r: r, c: c });
            if (sheet[addr] && sheet[addr].t !== 'z') {
                sheet[addr].s = { font: { color: { rgb: 'C00000' } } };
            }
        }
    }
}

// ============================================================
//  统计一组试次的结果行：两种超时口径的正确率与容量 K
//  durationLabel：该组对应的呈现时长（ms）；传「汇总」表示跨全部时长
//  K = (Acc * N^2 - N) / (N - 1)，Acc 为 0–1 小数
//
//  Sheet2 同时输出两种预先约定的口径（列名 2026-10-22 起统一为短后缀）：
//    Accuracy_Exclude / K_Exclude        ：超时从分母和正确数中剔除（主口径）；
//    Accuracy_TimeoutErr / K_TimeoutErr  ：超时保留在分母，按错误计入（备选口径）。
//  Trial Data 仍用 TrialStatus 与 CountedInK 明确标记超时，TotalTrials、
//  TimeoutTrials、ValidTrials 三列用于核对两种口径的样本数。
// ============================================================
const EXCLUDE_TIMEOUT_FROM_K = true;

function isCountedTrial(d) {
    // 缺列的历史数据（老版本导出的行）按"计入"处理，保证兼容
    if (EXCLUDE_TIMEOUT_FROM_K && (d.CountedInK === 0 || d.timeout === true)) return false;
    return true;
}

function buildKRow(ss, durationLabel, rows) {
    const validRows = rows.filter(isCountedTrial);
    const nValid = validRows.length;
    const nCorrect = validRows.filter(d => d.accuracy === 1).length;
    const accExclude = nValid > 0 ? nCorrect / nValid : 0;
    const nTotal = rows.length;
    const accTimeoutError = nTotal > 0 ? nCorrect / nTotal : 0;

    const correctRTs = validRows.filter(d => d.accuracy === 1 && typeof d.rt === 'number').map(d => d.rt);
    const rtMean = correctRTs.length > 0
        ? correctRTs.reduce((s, v) => s + v, 0) / correctRTs.length
        : '';

    const N = ss;
    const kExclude = N > 1 ? (accExclude * N * N - N) / (N - 1) : '';
    const kTimeoutError = N > 1 ? (accTimeoutError * N * N - N) / (N - 1) : '';

    // 低正确率标红：仅针对「set size ≤ 4 且 呈现时长 ≥ 200 ms」的明细条件行，
    // 且该 set size 在低正确率判定中确实触发了（≥200ms 试次正确率 < 50%）。
    // 汇总行（durationLabel === '汇总'）与 K 为负的条件行都不标红，避免把时长效应混入。
    const isDetailRow = durationLabel !== '汇总' && Number(durationLabel) >= LOW_ACC_MIN_DURATION_MS;
    const setSizeInRule = ss <= LOW_ACC_MIN_SET_SIZE;
    const ssDetail = lowAccDetail.details.find(d => d.setSize === ss);
    const lowAccRed = !!(isDetailRow && setSizeInRule && ssDetail && ssDetail.low && ss4LowAccuracyFlag);

    // 跨时长汇总行不需要 StudyDuration（该列在 K Values 中必须恒为数值）
    const isSummary = String(durationLabel) === '汇总';

    return {
        SubjectID: p.subjectID,
        SessionID: SESSION_ID,
        ExperimentVersion: EXPERIMENT_VERSION,
        ExperimentType: EXPERIMENT_NAME,
        subGroup: exportNum(p.subGroup),
        Order: exportNum(p.subGroup),
        SetSize: ss,
        StudyDuration: durationLabel,
        Condition: isSummary ? ('SS' + ss + '-汇总') : conditionLabel(durationLabel, ss),
        TotalTrials: nTotal,
        ValidTrials: nValid,
        TimeoutTrials: nTotal - nValid,
        TimeoutExclusionRate: (nTotal > 0 ? (nTotal - nValid) / nTotal : 0).toFixed(4),
        CorrectTrials: nCorrect,
        Accuracy_Exclude: accExclude.toFixed(4),
        K_Exclude: typeof kExclude === 'number' ? kExclude.toFixed(4) : '',
        Accuracy_TimeoutErr: accTimeoutError.toFixed(4),
        K_TimeoutErr: typeof kTimeoutError === 'number' ? kTimeoutError.toFixed(4) : '',
        MeanRT: rtMean === '' ? '' : rtMean.toFixed(4),
        LowAccFlag: ss4LowAccuracyFlag ? 1 : 0,
        LowAccRed: lowAccRed ? 1 : 0,
        LowAccBlockCount: lowAccBlockCount
    };
}

// ============================================================
//  按「呈现时长 × set size」分组重排试次
//  分组顺序：呈现时长升序 → set size 升序；组内保持原始实验顺序（不再排序）
//  行首附加 GroupIndex（分组序号，1 起）与 GroupLabel（如 50ms-SS3）两列
//  末尾兜底追加无法归类的行，确保不丢数据
// ============================================================
function groupTrialsByCondition(rows) {
    const durations = prefs.studyDurations.slice().sort((a, b) => a - b);
    const setSizes = prefs.setSizes.slice().sort((a, b) => a - b);

    const grouped = [];
    const placed = new Set();
    let groupIndex = 0;

    for (const duration of durations) {
        for (const ss of setSizes) {
            groupIndex++;
            const groupLabel = duration + 'ms-SS' + ss;
            for (const d of rows) {
                // 用 Number() 比较：字段恒为数字，但避免个别行被写成字符串时漏掉
                if (Number(d.studyDuration) === duration && Number(d.setSize) === ss) {
                    grouped.push(Object.assign({ GroupIndex: groupIndex, GroupLabel: groupLabel }, d));
                    placed.add(d);
                }
            }
        }
    }

    // 兜底：条件无法归类的行（正常不应发生）追加在末尾，避免静默丢数据
    for (const d of rows) {
        if (!placed.has(d)) {
            grouped.push(Object.assign({ GroupIndex: 0, GroupLabel: '未归类' }, d));
        }
    }
    return grouped;
}

// ============================================================
//  导出列规范（2026-10-22，为 R / SPSS 的 GLM / 混合模型而设）
//  ------------------------------------------------------------
//  四条硬规则——以后增删列时必须同步维护，否则统计端会踩坑：
//   ① 长格式（long format）：试次表一行一试次、条件表一行一条件。
//      R 的 lmer/aov、SPSS 的 MIXED 直接吃；不要再手工转宽表。
//   ② 自变量纯数值：SetSize / StudyDuration / Accuracy 等一律写数字；
//      文本只保留在 SubjectID / Condition / TrialStatus 这类标签列。
//      （数值列里混进一个 '汇总'，R 会把整列读成 character。）
//   ③ 缺失留空而不是写 0：超时试次的 RT / Response 写成空单元格，
//      R 读入为 NA、SPSS 读入为缺失值，不会被当成 0 秒 / 0 号反应。
//   ④ 列名纯 ASCII、无空格与中文：避免 R 的 make.names() 自动改名
//      导致脚本里的列名与 Excel 里看到的对不上。
//
//  列顺序（试次表）：标识 → 设计自变量 → 因变量 → 试次状态 → 诊断列
//                    → 被试信息 → 刺激细节（分析时可整块忽略）
// ============================================================

// 转成数值；非数值（空值 / '汇总' / 非数字字符串）一律返回 null → 导出为空单元格
function exportNum(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
}

// 条件标签：50ms-SS3 这类人眼可读的分组名（SPSS 里可直接当因子用）
function conditionLabel(duration, setSize) {
    return duration + 'ms-SS' + setSize;
}

// 按固定列规格生成 sheet：列顺序完全由 columns 决定，与对象键顺序无关
function sheetFromColumns(rows, columns) {
    const headers = columns.map(col => col.header);
    const data = rows.map((row, i) => {
        const out = {};
        for (const col of columns) {
            const v = col.get(row, i);
            if (v === null || v === undefined || v === '') continue;  // 空单元格 → NA
            out[col.header] = col.text ? String(v) : v;
        }
        return out;
    });
    // 显式传 header：即使首行某列缺值，列也不会被挤到后面或消失
    return XLSX.utils.json_to_sheet(data, { header: headers });
}

// --- Sheet 1 / Sheet 4（Trial Data、Grouped Trials）共用 ---
const TRIAL_EXPORT_COLUMNS = [
    { header: 'SubjectID',          text: true, get: (d) => d.subjectID },
    { header: 'SessionID',          text: true, get: (d) => d.sessionID || SESSION_ID },
    { header: 'ExperimentVersion',  text: true, get: (d) => d.experimentVersion || EXPERIMENT_VERSION },
    { header: 'ExperimentType',    text: true, get: (d) => d.experimentType || EXPERIMENT_NAME },
    // subGroup / Order：组别（0/1/2），0=调试，1/2=正式顺序，被试间变量。
    { header: 'subGroup',                   get: (d) => exportNum(d.subGroup !== undefined ? d.subGroup : p.subGroup) },
    { header: 'Order',                      get: (d) => exportNum(d.subGroup !== undefined ? d.subGroup : p.subGroup) },
    { header: 'Block',                      get: (d) => exportNum(d.block) },
    { header: 'TrialInBlock',               get: (d) => exportNum(d.trial) },
    { header: 'TrialGlobal',                get: (d) => exportNum(d.TrialGlobal) },
    { header: 'SetSize',                    get: (d) => exportNum(d.setSize) },
    { header: 'StudyDuration',              get: (d) => exportNum(d.studyDuration) },
    { header: 'Condition',         text: true, get: (d) => conditionLabel(d.studyDuration, d.setSize) },
    { header: 'Accuracy',                   get: (d) => exportNum(d.accuracy) },
    { header: 'RT',                         get: (d) => (typeof d.rt === 'number' ? Number(d.rt.toFixed(4)) : null) },
    { header: 'ResponseRGB',         text: true, get: (d) => d.responseRGB },
    { header: 'CorrectRGB',          text: true, get: (d) => d.correctRGB },
    { header: 'ResponseColorIdx',           get: (d) => exportNum(d.responseColorIdx) },
    { header: 'CorrectColorIdx',            get: (d) => exportNum(d.correctColorIdx) },
    { header: 'TrialStatus',        text: true, get: (d) => d.TrialStatus || (d.timeout ? 'timeout' : 'valid') },
    { header: 'CountedInK',                 get: (d) => exportNum(d.CountedInK) },
    { header: 'StudyDurationActual',        get: (d) => exportNum(d.studyDurationActual) },
    { header: 'RandomSeed',           text: true, get: (d) => d.rndSeed !== undefined ? d.rndSeed : p.rndSeed },
    { header: 'ResumedFromBlock',           get: (d) => exportNum(d.resumedFromBlock) },
    { header: 'ResumeCount',                get: (d) => exportNum(d.resumeCount) },
    { header: 'ChangeItem',                 get: (d) => exportNum(d.changeItem) },
    { header: 'OrigColorIdx',               get: (d) => exportNum(d.origColorIdx) },
    { header: 'ProbeColorIdx',              get: (d) => exportNum(d.probeColorIdx) },
    { header: 'ItemLocs',            text: true, get: (d) => d.itemLocs },
    { header: 'ItemColors',          text: true, get: (d) => d.itemColors },
    { header: 'LabelMap',           text: true, get: (d) => d.labelMap },
    { header: 'TrialLabels',         text: true, get: (d) => d.trialLabels },
    { header: 'LabelChars',          text: true, get: (d) => d.labelChars }
];

// --- Sheet 2（K Values）---
// 列名即 buildKRow 返回的字段名，此处只负责「固定顺序 + 统一类型」
const K_EXPORT_COLUMNS = [
    'SubjectID', 'SessionID', 'ExperimentVersion', 'ExperimentType', 'subGroup', 'Order',
    'SetSize', 'StudyDuration', 'Condition',
    'TotalTrials', 'ValidTrials', 'TimeoutTrials', 'TimeoutExclusionRate',
    'CorrectTrials', 'Accuracy_Exclude', 'K_Exclude',
    'Accuracy_TimeoutErr', 'K_TimeoutErr', 'MeanRT',
    'LowAccFlag', 'LowAccRed', 'LowAccBlockCount'
];
const K_TEXT_COLUMNS = new Set(['SubjectID', 'ExperimentType', 'Condition']);

function kSheetFromRows(rows) {
    const data = rows.map(r => {
        const out = {};
        for (const header of K_EXPORT_COLUMNS) {
            const v = r[header];
            if (v === null || v === undefined || v === '') continue;
            if (K_TEXT_COLUMNS.has(header)) { out[header] = String(v); continue; }
            const n = Number(v);
            // 非数值（如汇总行的 StudyDuration='汇总'）自动留空，
            // 保证 StudyDuration 在 K Values 里恒为数值自变量
            if (Number.isFinite(n)) out[header] = n;
        }
        return out;
    });
    return XLSX.utils.json_to_sheet(data, { header: K_EXPORT_COLUMNS });
}

function downloadXLSX() {
    // 计算低正确率标记（SS3/SS4 在 ≥200ms 正确率 < 50%）
    const lowAccResult = computeLowAccuracyFlag();
    ss4LowAccuracyFlag = lowAccResult.triggered;
    lowAccDetail = lowAccResult;

    // 补全被试信息与断点续做标记（幂等：finishExperiment 里可能已补过一次），
    // 保证 Trial Data 每行都带 subGroup / subjectID / RandomSeed 等关键字段
    enrichRows(stimData);

    // 全局试次序号（1..360）：长格式里用来查顺序效应 / 疲劳效应，
    // 先按原始实验顺序写好，Grouped Trials 重排后仍指向同一试次
    stimData.forEach((d, i) => { d.TrialGlobal = i + 1; });

    // ============================================================
    //  Sheet 1: Trial Data
    //  长格式，一行一试次；列顺序与含义见上方「导出列规范」
    // ============================================================
    const trialSheet = sheetFromColumns(stimData, TRIAL_EXPORT_COLUMNS);

    // ============================================================
    //  Sheet 2: K Values（条件级）
    //  CLT 公式: K = (Acc * N^2 - N) / (N - 1)
    //  同时报告两种口径：
    //    Accuracy_Exclude / K_Exclude        ：超时剔除（主口径）；
    //    Accuracy_TimeoutErr / K_TimeoutErr ：超时按错误计入（备选口径）。
    //  TimeoutExclusionRate 报告每个条件被剔除的超时比例。
    //  StudyDuration 在本表恒为数值；跨时长的汇总行单独放 Sheet 4。
    // ============================================================
    const kData = [];
    const sortedDurations = prefs.studyDurations.slice().sort((a, b) => a - b);
    const sortedSetSizes = prefs.setSizes.slice().sort((a, b) => a - b);
    for (const duration of sortedDurations) {
        for (const ss of sortedSetSizes) {
            const rows = stimData.filter(d =>
                Number(d.setSize) === ss && Number(d.studyDuration) === duration);
            kData.push(buildKRow(ss, duration, rows));
        }
    }
    const kSheet = kSheetFromRows(kData);
    // 标红低正确率条件行（XLSX 0.18.5 无样式能力时为 no-op，见函数注释）
    markKRowsRed(kSheet);

    // ============================================================
    //  Sheet 3: ColorAccuracy（每个颜色在每个 set size 下的反应正确率）
    //  行：颜色 RGB；列：set size；值：该颜色作为正确选项时的正确率
    // ============================================================
    const colorAccData = [];
    for (const ss of sortedSetSizes) {
        const ssRows = stimData.filter(d => Number(d.setSize) === ss);
        for (let c = 0; c < colors_9.length; c++) {
            const colorRows = ssRows.filter(d => Number(d.correctColorIdx) === c);
            const nTotal = colorRows.length;
            const nCorrect = colorRows.filter(d => d.accuracy === 1).length;
            const acc = nTotal > 0 ? nCorrect / nTotal : null;
            colorAccData.push({
                SetSize: ss,
                ColorIdx: c,
                RGB: colors_9[c].join(','),
                TotalTrials: nTotal,
                CorrectTrials: nCorrect,
                Accuracy: acc
            });
        }
    }
    const colorAccSheet = XLSX.utils.json_to_sheet(colorAccData, {
        header: ['SetSize', 'ColorIdx', 'RGB', 'TotalTrials', 'CorrectTrials', 'Accuracy']
    });

    // ============================================================
    //  Sheet 4: Grouped Trials
    //  列与 Sheet1 完全相同（前面多 GroupIndex / GroupLabel 两列标识分组），
    //  但按「呈现时长 → set size」分组重排：同一条件的试次在表中连续排列，
    //  组内仍保持原始实验顺序，不做二次排序。
    // ============================================================
    const groupedColumns = [
        { header: 'GroupIndex', get: (d) => exportNum(d.GroupIndex) },
        { header: 'GroupLabel', text: true, get: (d) => d.GroupLabel }
    ].concat(TRIAL_EXPORT_COLUMNS);
    const groupedSheet = sheetFromColumns(groupTrialsByCondition(stimData), groupedColumns);

    // === 生成 Excel 文件 ===
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, trialSheet, 'Trial Data');
    XLSX.utils.book_append_sheet(wb, kSheet, 'K Values');
    XLSX.utils.book_append_sheet(wb, colorAccSheet, 'ColorAccuracy');
    XLSX.utils.book_append_sheet(wb, groupedSheet, 'Grouped Trials');

    // 文件名规则：实验名称_被试编号（如 exp1_clt_compare_1_023_abc.xlsx）
    // 存在低正确率条件（SS3/SS4 在 ≥200ms 正确率 < 50%）时在最前面加 F 标记；
    // 调试模式不做该判定，因此永不带 F
    const filePrefix = (!p.debugMode && ss4LowAccuracyFlag) ? 'F' : '';
    const fileName = filePrefix + EXPERIMENT_NAME + '_' + p.subjectID + '.xlsx';
    console.log(`[exp1_clt] 导出模式=${p.debugMode ? '调试' : '正式'} 被试编号=${p.subjectID} 文件名=${fileName}`);
    saveWorkbookFile(wb, fileName);
}

// ============================================================
//  写出 Excel 文件
//  用正确的 xlsx MIME + 明确的文件名触发下载；同时在结束界面保留一个
//  手动下载链接（某些浏览器/内嵌预览不遵循 download 属性，会把文件存成
//  随机名或没有扩展名，看起来「不是 Excel 文件」，此时可用该链接另存）
// ============================================================
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
let lastExport = null;   // { url, fileName, size }

function saveWorkbookFile(wb, fileName) {
    if (typeof XLSX === 'undefined') {
        throw new Error('Excel 导出库 (xlsx) 未加载');
    }
    const bytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const u8 = (bytes instanceof Uint8Array) ? bytes : new Uint8Array(bytes);
    // 校验 zip 魔数 'PK'，确保写出的是合法 xlsx（而不是被当成其它格式）
    if (u8[0] !== 0x50 || u8[1] !== 0x4B) {
        throw new Error('生成的 Excel 内容异常');
    }

    const blob = new Blob([u8], { type: XLSX_MIME });
    const url = URL.createObjectURL(blob);
    lastExport = { url: url, fileName: fileName, size: u8.length };

    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.rel = 'noopener';
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();

    console.log(`[exp1_clt] 已导出 ${fileName}（${u8.length} 字节，${wb.SheetNames.length} 个 sheet）`);
    showManualDownload(fileName, url);
}

// 结束界面上的手动下载入口（文件名与自动下载一致）
function showManualDownload(fileName, url) {
    const box = document.getElementById('manual-download');
    if (!box) return;
    const link = document.getElementById('manual-download-link');
    if (link) {
        link.href = url;
        link.download = fileName;
        link.innerText = fileName;
    }
    box.classList.remove('hidden');
}

// ============================================================
//  Canvas 绘图工具
// ============================================================
function drawBackground() {
    ctx.fillStyle = 'rgb(128,128,128)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
}

function drawFixation() {
    ctx.fillStyle = 'rgb(0,0,0)';
    ctx.beginPath();
    ctx.arc(canvas.width / 2, canvas.height / 2, prefs.fixationSize, 0, Math.PI * 2);
    ctx.fill();
}

function drawSquare(x, y, colorArr) {
    ctx.fillStyle = `rgb(${colorArr[0]},${colorArr[1]},${colorArr[2]})`;
    const half = prefs.stimSize / 2;
    ctx.fillRect(x - half, y - half, prefs.stimSize, prefs.stimSize);
}

// compare 版自绘虚拟光标：位置算法保持随机圆内生成，交互机制与 circle 版一致。
function drawVirtualCursor() {
    if (!cursorActive) return;
    const r = 9, pad = r + 2;
    const x = Math.max(pad, Math.min(canvas.width - pad, virtualCursor.x));
    const y = Math.max(pad, Math.min(canvas.height - pad, virtualCursor.y));
    ctx.save();
    ctx.beginPath(); ctx.arc(x, y, r, 0, Math.PI * 2);
    ctx.lineWidth = 3.5; ctx.strokeStyle = 'rgb(255,255,255)'; ctx.stroke();
    ctx.beginPath(); ctx.arc(x, y, r, 0, Math.PI * 2);
    ctx.lineWidth = 1.5; ctx.strokeStyle = 'rgb(0,0,0)'; ctx.stroke();
    ctx.restore();
}

function redrawProbeScreen() {
    if (!probeSnapshot) return;
    drawBackground(); drawFixation();
    for (let i = 0; i < probeSnapshot.locs.length; i++) {
        drawSquare(probeSnapshot.locs[i].x, probeSnapshot.locs[i].y,
            colors_9[probeSnapshot.colorIdx[i]]);
    }
    drawVirtualCursor();
}

function resetVirtualCursor() {
    virtualCursor.x = canvas.width / 2;
    virtualCursor.y = canvas.height / 2;
    cursorStartX = virtualCursor.x; cursorStartY = virtualCursor.y;
    cursorTravelPx = 0; lastClientX = null; lastClientY = null;
    cursorSkipNextMove = false; cursorActive = true;
    canvas.style.cursor = 'none';
}

function hideVirtualCursor() {
    cursorActive = false; probeSnapshot = null;
    canvas.style.cursor = 'default';
}

window.addEventListener('mousemove', (e) => {
    if (!cursorActive) return;
    let dx = 0, dy = 0;
    if (pointerLocked && typeof e.movementX === 'number' && typeof e.movementY === 'number') {
        dx = e.movementX; dy = e.movementY;
    } else if (lastClientX !== null && Number.isFinite(e.clientX) && Number.isFinite(e.clientY)) {
        dx = e.clientX - lastClientX; dy = e.clientY - lastClientY;
    }
    lastClientX = e.clientX; lastClientY = e.clientY;
    if (cursorSkipNextMove) { cursorSkipNextMove = false; return; }
    if (dx === 0 && dy === 0) return;
    const rect = canvas.getBoundingClientRect();
    const sx = rect.width ? canvas.width / rect.width : 1;
    const sy = rect.height ? canvas.height / rect.height : 1;
    const mx = dx * sx, my = dy * sy;
    virtualCursor.x += mx; virtualCursor.y += my;
    cursorTravelPx += Math.hypot(mx, my);
    redrawProbeScreen();
});

document.addEventListener('pointerlockchange', () => {
    pointerLocked = document.pointerLockElement === canvas;
    lastClientX = null; lastClientY = null;
    if (!pointerLocked && cursorActive &&
        (experimentPhase === 'trial_response' || experimentPhase === 'practice_response')) {
        hideVirtualCursor();
        alert('鼠标指针锁定已退出，当前试次已停止。请重新进入实验。');
    }
});

document.addEventListener('fullscreenchange', () => {
    if (!document.fullscreenElement && cursorActive &&
        (experimentPhase === 'trial_response' || experimentPhase === 'practice_response')) {
        hideVirtualCursor();
        if (document.pointerLockElement === canvas && document.exitPointerLock) document.exitPointerLock();
        alert('全屏已退出，当前试次已停止。请重新进入实验。');
    }
});

// ============================================================
//  getStimLocs — 圆环随机位置（compare 版 CDT / CLT 统一布局）
//
//  参考 Luck & Vogel (1997) 与 2021 AP&P 圆形阵列：所有刺激中心落在
//  同一个以注视点为圆心的圆周上，位置等角间隔、逐试次随机旋转。
//
//  ① 圆环（"位置都必须在圆环上"）：
//     半径 R 对所有色块恒定，每个色块到注视点距离完全相同，
//     不会出现某个色块离注视点"过远"，也不会被挤到画布外"过近"。
//     名义半径按画布短边缩放（768px 参考系下 200px ≈ 1.95° 视角），
//     再内缩 半个色块 + 留白，保证 51px 色块完整可见。
//
//  ② 最小间距（"间距按文献设置 / 没有过近"）：
//     圆周等分为 RING_SLOTS 个候选槽位，相邻槽位中心距
//     = 2R·sin(π / RING_SLOTS)。把相邻色块最小中心距设为
//     MIN_CENTER_DIST = 色块边长 + RING_MIN_GAP，反推所需最小半径
//     R_min = MIN_CENTER_DIST / (2·sin(π / RING_SLOTS))，
//     取 R = max(R_名义, R_min)，从几何上保证任意相邻色块不重叠、不贴边。
//     ※ 需按具体文献调整间距时，只改 RING_SLOTS / RING_MIN_GAP 两个参数。
//
//  ③ 位置随机 + 分布均匀（"位置选择随机 / 没有过近或过远"）：
//     每试次给整圈一个随机旋转 φ ∈ (−半槽距, +半槽距)，幅度不超过
//     半槽距，槽位不会跨过坐标轴线；再在等间隔槽位里随机挑 setSize 个，
//     用"环形相邻间隔的均衡度"给组合评分，只从最均匀的一批里选，
//     避免色块全挤在半圈（一侧留大片空白"过远"）或扎堆（"过近"）。
//
//  ④ 降低重复（"相同位置选择的试次较少"）：
//     记录上一次使用的槽位组合签名，随机抽样时尽量避开重复组合；
//     叠加连续随机旋转角，使连续试次几乎不会呈现完全相同的布局。
// ============================================================
const RING_SLOTS = 8;                    // 圆周等分候选位置数（文献常用 8）
const RING_SLOT_STEP = 360 / RING_SLOTS; // 槽位角间隔 = 45°
const RING_MIN_GAP = 12;                 // 相邻色块最小额外留白 px
const RING_EDGE_MARGIN = 8;              // 色块中心到画布边缘的最小留白 px
const RING_RADIUS_RATIO = 200 / 768;     // 名义半径 / 画布短边（≈1.95° 视角）
let lastRingSignature = null;            // 上一次使用的槽位组合，用于降低重复

function getStimLocs(nItems, cx, cy) {
    const S = prefs.stimSize;
    const shortSide = Math.min(canvas.width, canvas.height);
    const nominalR = Math.min(
        shortSide * RING_RADIUS_RATIO,
        shortSide / 2 - S / 2 - RING_EDGE_MARGIN
    );
    const minCenterDist = S + RING_MIN_GAP;
    const minR = minCenterDist / (2 * Math.sin(Math.PI / RING_SLOTS));
    const radius = Math.max(nominalR, minR);

    // 枚举从 RING_SLOTS 个槽位中选 nItems 个的所有组合。
    // 用标称角（不含旋转）判断左右半区；要求左右各至少一个色块，
    // 保证变化觉察 / 变化定位任务左右半区都有可点色块。
    const baseDeg = RING_SLOT_STEP / 2;   // 22.5°，避开水平/垂直坐标轴
    const combos = [];
    const collect = (start, picked) => {
        if (picked.length === nItems) {
            let left = 0;
            for (const i of picked) {
                const nominalA = (baseDeg + i * RING_SLOT_STEP) * Math.PI / 180;
                if (Math.cos(nominalA) < 0) left++;
            }
            if (left < 1 || nItems - left < 1) return;
            const gaps = [];
            for (let k = 0; k < nItems; k++) {
                const a = picked[k];
                const b = k + 1 < nItems ? picked[k + 1] : picked[0] + RING_SLOTS;
                gaps.push(b - a);
            }
            const mean = RING_SLOTS / nItems;
            const uneven = gaps.reduce((s, g) => s + (g - mean) ** 2, 0);
            combos.push({ indices: picked.slice(), uneven, sig: picked.join(',') });
            return;
        }
        for (let i = start; i < RING_SLOTS; i++) {
            picked.push(i); collect(i + 1, picked); picked.pop();
        }
    };
    collect(0, []);

    // 只保留最均匀的一批组合，再随机挑选，且尽量避开上一次的组合。
    const bestUneven = Math.min(...combos.map(c => c.uneven));
    let pool = combos.filter(c => c.uneven === bestUneven);
    if (pool.length > 1 && lastRingSignature !== null) {
        const fresh = pool.filter(c => c.sig !== lastRingSignature);
        if (fresh.length) pool = fresh;
    }
    const chosen = pool[randInt(0, pool.length - 1)];
    lastRingSignature = chosen.sig;

    // 整圈随机旋转：幅度小于半槽距，槽位不跨轴线；使绝对角度逐试次变化。
    const rotation = (random() - 0.5) * RING_SLOT_STEP * Math.PI / 180;

    const picked = chosen.indices.slice();
    shuffleArray(picked);
    return picked.map(idx => {
        const a = (baseDeg + idx * RING_SLOT_STEP) * Math.PI / 180 + rotation;
        return { x: cx + radius * Math.cos(a), y: cy + radius * Math.sin(a) };
    });
}

