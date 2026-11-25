#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
exp1_clt_compare_psychopy.py
变化定位任务 (Change Localization Task, CLT) - PsychoPy 版本
基于 clt_compare 实验规则移植

实验设计：
  - set size: 4 / 6
  - 呈现时长: 250 ms（compare 版本单一条件）
  - 每 set size 40 trial，共 80 trial
  - 4 个 block，每 block 20 trial
  - block 间强制休息 15 秒（倒计时 + 按空格继续）
  - 练习 16 trial（每 set size 各 8 trial）
  - 反应时限 4000 ms

视觉参数（与 JS 版一致）：
  - 背景色：rgb(128,128,128) 浅灰色
  - 注视点：黑色圆点，直径 16px（半径 8px）
  - 色块大小：51x51 px
  - 色块最小间距：127.5 px
  - 全屏运行
  - 虚拟光标：白色圆圈 + 黑色边框，默认位于屏幕中央
  - 隐藏系统光标
"""

import os
import sys
import json
import math
import random
import datetime
import numpy as np
from psychopy import visual, core, event, gui, data, logging
from psychopy.constants import NOT_STARTED
import pandas as pd

# 关闭 PsychoPy 默认的过多日志
logging.console.setLevel(logging.WARNING)

# ============================================================
#  实验参数
# ============================================================
EXPERIMENT_NAME = 'exp1_clt_compare'
EXPERIMENT_VERSION = 'clt_compare'
SET_SIZES = [4, 6]
STUDY_DURATIONS = [250]  # ms，compare 版本单一条件
NUM_BLOCKS = 4
TRIALS_PER_BLOCK = 20
N_PER_SET_SIZE = 40  # 每 set size 40 trial
ITI_MS = 500
RETENTION_INTERVAL_MS = 1000
RESPONSE_LIMIT_MS = 4000
BREAK_MIN_SECONDS = 15
STIM_SIZE_PX = 51
FIXATION_RADIUS_PX = 8
LABEL_FONT_SIZE = 18
N_PRACTICE_PER_SET = 8  # 每 set size 练习 8 trial

# 9 色调色板（RGB 0-255）
COLORS_RGB = [
    (255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0),
    (255, 0, 255), (0, 255, 255), (255, 255, 255), (0, 0, 0), (255, 128, 0)
]

# ============================================================
#  被试信息对话框
# ============================================================
def get_subject_info(win):
    """弹出对话框收集被试信息，并显示自动生成的 SubjectID"""
    dlg = gui.Dlg(title='实验设置')
    dlg.addField('组别 (0=调试, 1=先FCDT后CLT, 2=先CLT后FCDT):', choices=['1', '2', '0'])
    dlg.addField('Subject ID (留空自动生成):', '')
    
    dlg.show()
    if not dlg.OK:
        core.quit()
    
    sub_group = int(dlg.data[0])
    subject_id = dlg.data[1].strip()
    
    # 自动生成 SubjectID: {组别}_{3位点击顺序}_{3位随机后缀}
    if not subject_id:
        click_order = random.randint(0, 999)
        suffix = ''.join(random.choices('abcdefghijklmnopqrstuvwxyz0123456789', k=3))
        subject_id = f"{sub_group}_{click_order:03d}_{suffix}"
    
    # 调试模式固定编号
    if sub_group == 0:
        subject_id = '1111'
    
    # 显示 SubjectID 确认页面
    info_text = visual.TextStim(win, 
        text=f'你的被试编号是：\n\n{subject_id}\n\n请记下这个编号，然后按空格键继续',
        color='black', height=32, wrapWidth=600)
    
    info_text.draw()
    win.flip()
    event.waitKeys(keyList=['space'])
    
    return {
        'subGroup': sub_group,
        'debugMode': sub_group == 0,
        'subjectID': subject_id,
        'rndSeed': 1111 if sub_group == 0 else hash(subject_id) % 2**31,
        'expDate': datetime.datetime.now().strftime('%Y%m%d')
    }

# ============================================================
#  PRNG（与 JS 版一致）
# ============================================================
class SeededRandom:
    def __init__(self, seed):
        self.seed = seed
    
    def random(self):
        x = math.sin(self.seed) * 10000
        self.seed += 1
        return x - math.floor(x)
    
    def randint(self, a, b):
        return math.floor(self.random() * (b - a + 1)) + a
    
    def shuffle(self, arr):
        for i in range(len(arr) - 1, 0, -1):
            j = self.randint(0, i)
            arr[i], arr[j] = arr[j], arr[i]
        return arr

# ============================================================
#  刺激位置生成（getStimLocs 移植）
# ============================================================
def get_stim_locs(n_items, cx, cy, screen_size, rng):
    """
    生成刺激位置（极坐标四象限）
    返回相对于屏幕中心 (cx, cy) 的偏移量列表，每个元素为 {'x': ..., 'y': ...}
    """
    S = STIM_SIZE_PX
    short_side = min(screen_size)
    stimulus_circle_radius = short_side * (200.0 / 768.0)
    
    # 修复：确保 min_r < max_r
    max_r = min(stimulus_circle_radius, short_side / 2.0 - S)
    min_r = max(S * 1.5, max_r * 0.3)
    
    min_dist = S * 2.5
    axis_margin = 20 * math.pi / 180
    
    quad_patterns_6 = [
        [2, 2, 1, 1], [1, 2, 1, 2], [1, 1, 2, 2], [1, 2, 2, 1], [2, 1, 1, 2], [2, 1, 2, 1]
    ]
    
    if n_items == 3:
        per_quad = [1, 1, 1, 0]
        rng.shuffle(per_quad)
    elif n_items == 4:
        per_quad = [1, 1, 1, 1]
    elif n_items == 6:
        per_quad = quad_patterns_6[rng.randint(0, 5)]
    else:
        raise ValueError(f'nItems must be 3, 4, or 6, got {n_items}')
    
    q_lo = [
        axis_margin, math.pi / 2 + axis_margin,
        math.pi + axis_margin, 3 * math.pi / 2 + axis_margin
    ]
    q_hi = [
        math.pi / 2 - axis_margin, math.pi - axis_margin,
        3 * math.pi / 2 - axis_margin, 2 * math.pi - axis_margin
    ]
    
    n_candidates = 30
    best_score = -float('inf')
    best_x_pos = []
    best_y_pos = []
    
    for cand in range(n_candidates):
        x_pos_c = [0] * n_items
        y_pos_c = [0] * n_items
        idx = 0
        
        for q in range(4):
            k = per_quad[q]
            if k == 0:
                continue
            q_range = q_hi[q] - q_lo[q]
            
            for i in range(k):
                if k == 1:
                    theta_c = q_lo[q] + q_range * 0.5
                    jit = q_range * 0.15 * (rng.random() - 0.5) * 2
                    sub_lo = theta_c + jit - q_range * 0.05
                    sub_hi = theta_c + jit + q_range * 0.05
                else:
                    theta_c = q_lo[q] + q_range * (0.15 + 0.70 * i / (k - 1))
                    jit = q_range * 0.05 * (rng.random() - 0.5) * 2
                    sub_lo = theta_c + jit - q_range * 0.03
                    sub_hi = theta_c + jit + q_range * 0.03
                
                placed = False
                for att in range(500):
                    r = min_r + (max_r - min_r) * math.sqrt(rng.random())
                    theta = sub_lo + (sub_hi - sub_lo) * rng.random()
                    xi = cx + r * math.cos(theta)
                    yi = cy + r * math.sin(theta)
                    if abs(xi - cx) < S or abs(yi - cy) < S:
                        continue
                    too_close = False
                    for j in range(idx):
                        if math.hypot(xi - x_pos_c[j], yi - y_pos_c[j]) < min_dist:
                            too_close = True
                            break
                    if not too_close:
                        x_pos_c[idx] = xi
                        y_pos_c[idx] = yi
                        placed = True
                        break
                
                if not placed:
                    for att in range(800):
                        r = min_r + (max_r - min_r) * math.sqrt(rng.random())
                        theta = q_lo[q] + q_range * rng.random()
                        xi = cx + r * math.cos(theta)
                        yi = cy + r * math.sin(theta)
                        if abs(xi - cx) < S or abs(yi - cy) < S:
                            continue
                        too_close = False
                        for j in range(idx):
                            if math.hypot(xi - x_pos_c[j], yi - y_pos_c[j]) < min_dist:
                                too_close = True
                                break
                        if not too_close:
                            x_pos_c[idx] = xi
                            y_pos_c[idx] = yi
                            placed = True
                            break
                
                if not placed:
                    best_min_d = -1
                    best_xi = 0
                    best_yi = 0
                    for att in range(2000):
                        r = min_r + (max_r - min_r) * math.sqrt(rng.random())
                        theta = q_lo[q] + q_range * rng.random()
                        xi = cx + r * math.cos(theta)
                        yi = cy + r * math.sin(theta)
                        if abs(xi - cx) < S or abs(yi - cy) < S:
                            continue
                        min_d = float('inf')
                        for j in range(idx):
                            d = math.hypot(xi - x_pos_c[j], yi - y_pos_c[j])
                            if d < min_d:
                                min_d = d
                        if min_d > best_min_d:
                            best_min_d = min_d
                            best_xi = xi
                            best_yi = yi
                    x_pos_c[idx] = best_xi
                    y_pos_c[idx] = best_yi
                
                idx += 1
        
        # 评分
        dists = []
        for a in range(n_items - 1):
            for b in range(a + 1, n_items):
                dists.append(math.hypot(x_pos_c[a] - x_pos_c[b], x_pos_c[a] - x_pos_c[b]))
        
        angles = []
        for i in range(n_items):
            angles.append(math.atan2(y_pos_c[i] - cy, x_pos_c[i] - cx))
        angles = [(a + 2 * math.pi) % (2 * math.pi) for a in angles]
        angles.sort()
        angular_gaps = []
        for i in range(n_items):
            next_a = angles[i + 1] if i < n_items - 1 else angles[0] + 2 * math.pi
            angular_gaps.append(next_a - angles[i])
        
        min_d = min(dists)
        mean_d = sum(dists) / len(dists)
        std_d = math.sqrt(sum((d - mean_d) ** 2 for d in dists) / (len(dists) - 1)) if len(dists) > 1 else 0
        mean_g = sum(angular_gaps) / len(angular_gaps)
        std_g = math.sqrt(sum((g - mean_g) ** 2 for g in angular_gaps) / (len(angular_gaps) - 1)) if len(angular_gaps) > 1 else 0
        
        score = min_d - 0.4 * std_d - 40 * std_g
        if score > best_score:
            best_score = score
            best_x_pos = x_pos_c[:]
            best_y_pos = y_pos_c[:]
    
    # 转换为相对于 (cx, cy) 的偏移量，并适配 PsychoPy 坐标系统
    # JS 版：原点在左上角，Y 轴向下
    # PsychoPy：原点在中心，Y 轴向上
    locs = []
    for i in range(n_items):
        locs.append({
            'x': best_x_pos[i] - cx,
            'y': -(best_y_pos[i] - cy)  # 翻转 Y 轴
        })
    return locs

# ============================================================
#  试次生成
# ============================================================
def build_probe_plan(set_size, rng):
    """生成一次试次的完整视觉计划"""
    color_idx_arr = list(range(9))
    rng.shuffle(color_idx_arr)
    mem_colors = color_idx_arr[:set_size]
    
    # 变化方块：随机 1 个
    change_idx = rng.randint(0, set_size - 1)
    
    # 新颜色：排除记忆阵列出现过的全部颜色
    new_color_pool = [c for c in range(9) if c not in mem_colors]
    if not new_color_pool:
        raise ValueError('build_probe_plan: 调色板颜色不足')
    new_color = new_color_pool[rng.randint(0, len(new_color_pool) - 1)]
    
    # 探测阵列颜色：仅替换变化方块
    probe_colors = mem_colors[:]
    probe_colors[change_idx] = new_color
    
    # 内部标签映射
    lmap = list(range(set_size))
    rng.shuffle(lmap)
    
    correct_label = lmap.index(change_idx) + 1
    
    return {
        'setSize': set_size,
        'memColors': mem_colors,
        'probeColors': probe_colors,
        'changeIdx': change_idx,
        'origColor': mem_colors[change_idx],
        'newColor': new_color,
        'labelMap': lmap,
        'correctLabel': correct_label
    }

# ============================================================
#  K 值计算
# ============================================================
def compute_k(accuracy, n):
    """CLT 公式: K = (Acc * N^2 - N) / (N - 1)"""
    if n <= 1:
        return None
    return (accuracy * n * n - n) / (n - 1)

def compute_low_acc_flag(trials):
    """低正确率判定：set size 3/4 在 ≥200ms 条件下正确率 < 50%"""
    details = []
    for ss in [3, 4]:
        rows = [t for t in trials if t['setSize'] == ss and t['studyDuration'] >= 200]
        n = len(rows)
        if n == 0:
            details.append({'setSize': ss, 'accuracyPct': None, 'nTotal': 0, 'low': False})
            continue
        acc = sum(t['accuracy'] for t in rows) / n
        acc_pct = acc * 100
        details.append({
            'setSize': ss,
            'accuracyPct': acc_pct,
            'nTotal': n,
            'low': acc_pct < 50
        })
    triggered = any(d['low'] for d in details)
    return {'triggered': triggered, 'details': details}

# ============================================================
#  数据导出
# ============================================================
def export_to_excel(all_trials, subject_info, file_prefix=''):
    """导出 Excel（Trial Data / K Values / ColorAccuracy / Grouped Trials）"""
    import openpyxl
    from openpyxl import Workbook
    
    subject_id = subject_info['subjectID']
    sub_group = subject_info['subGroup']
    debug_mode = subject_info['debugMode']
    
    # 计算低正确率标记
    low_acc = compute_low_acc_flag(all_trials)
    low_acc_flag = low_acc['triggered'] and not debug_mode
    
    # 文件名
    file_name = f"{file_prefix}{EXPERIMENT_NAME}_{subject_id}.xlsx"
    
    wb = Workbook()
    
    # ==================== Sheet 1: Trial Data ====================
    ws1 = wb.active
    ws1.title = 'Trial Data'
    
    trial_headers = [
        'SubjectID', 'ExperimentVersion', 'ExperimentType', 'subGroup', 'Block', 'TrialInBlock',
        'TrialGlobal', 'SetSize', 'StudyDuration', 'Condition', 'Accuracy', 'RT',
        'ResponseRGB', 'CorrectRGB', 'ResponseColorIdx', 'CorrectColorIdx',
        'TrialStatus', 'CountedInK', 'StudyDurationActual', 'ChangeItem', 'OrigColorIdx', 'ProbeColorIdx'
    ]
    ws1.append(trial_headers)
    
    for i, t in enumerate(all_trials, 1):
        row = [
            t.get('subjectID', subject_id),
            EXPERIMENT_VERSION,
            EXPERIMENT_NAME,
            sub_group,
            t.get('block', ''),
            t.get('trial', ''),
            i,  # TrialGlobal
            t.get('setSize', ''),
            t.get('studyDuration', ''),
            f"{t.get('studyDuration', '')}ms-SS{t.get('setSize', '')}",
            t.get('accuracy', ''),
            t.get('rt', ''),
            t.get('responseRGB', ''),
            t.get('correctRGB', ''),
            t.get('responseColorIdx', ''),
            t.get('correctColorIdx', ''),
            t.get('trialStatus', 'valid' if not t.get('timeout') else 'timeout'),
            0 if t.get('timeout') else 1,
            t.get('studyDurationActual', ''),
            t.get('changeItem', ''),
            t.get('origColorIdx', ''),
            t.get('probeColorIdx', '')
        ]
        ws1.append(row)
    
    # ==================== Sheet 2: K Values ====================
    ws2 = wb.create_sheet('K Values')
    k_headers = [
        'SubjectID', 'ExperimentVersion', 'ExperimentType', 'subGroup', 'SetSize', 'StudyDuration',
        'Condition', 'TotalTrials', 'ValidTrials', 'TimeoutTrials', 'TimeoutExclusionRate',
        'CorrectTrials', 'Accuracy_Exclude', 'K_Exclude', 'Accuracy_TimeoutErr', 'K_TimeoutErr',
        'MeanRT', 'LowAccFlag', 'LowAccRed', 'LowAccBlockCount'
    ]
    ws2.append(k_headers)
    
    for ss in SET_SIZES:
        for dur in STUDY_DURATIONS:
            rows = [t for t in all_trials if t['setSize'] == ss and t['studyDuration'] == dur]
            total = len(rows)
            valid = [t for t in rows if not t.get('timeout')]
            n_valid = len(valid)
            n_correct = sum(1 for t in valid if t.get('accuracy') == 1)
            acc_exclude = n_correct / n_valid if n_valid > 0 else 0
            k_exclude = compute_k(acc_exclude, ss) if n_valid > 0 else None
            
            acc_timeout = n_correct / total if total > 0 else 0
            k_timeout = compute_k(acc_timeout, ss) if total > 0 else None
            
            correct_rts = [t['rt'] for t in valid if t.get('accuracy') == 1 and t.get('rt') is not None]
            mean_rt = sum(correct_rts) / len(correct_rts) if correct_rts else None
            
            timeout_rate = (total - n_valid) / total if total > 0 else 0
            
            # 低正确率判定
            is_detail = dur >= 200 and ss <= 4
            ss_detail = next((d for d in low_acc['details'] if d['setSize'] == ss), None)
            low_acc_red = 1 if (is_detail and ss_detail and ss_detail['low'] and low_acc['triggered']) else 0
            
            row = [
                subject_id, EXPERIMENT_VERSION, EXPERIMENT_NAME, sub_group, ss, dur,
                f"{dur}ms-SS{ss}",
                total, n_valid, total - n_valid, round(timeout_rate, 4),
                n_correct, round(acc_exclude, 4), round(k_exclude, 4) if k_exclude is not None else '',
                round(acc_timeout, 4), round(k_timeout, 4) if k_timeout is not None else '',
                round(mean_rt, 4) if mean_rt is not None else '',
                1 if low_acc_flag else 0,
                low_acc_red,
                0  # LowAccBlockCount
            ]
            ws2.append(row)
    
    # ==================== Sheet 3: ColorAccuracy ====================
    ws3 = wb.create_sheet('ColorAccuracy')
    ws3.append(['SetSize', 'ColorIdx', 'RGB', 'TotalTrials', 'CorrectTrials', 'Accuracy'])
    
    for ss in SET_SIZES:
        ss_rows = [t for t in all_trials if t['setSize'] == ss]
        for c in range(9):
            color_rows = [t for t in ss_rows if t.get('correctColorIdx') == c]
            n_total = len(color_rows)
            n_correct = sum(1 for t in color_rows if t.get('accuracy') == 1)
            acc = n_correct / n_total if n_total > 0 else None
            row = [
                ss, c, ','.join(map(str, COLORS_RGB[c])),
                n_total, n_correct,
                round(acc, 4) if acc is not None else ''
            ]
            ws3.append(row)
    
    # ==================== Sheet 4: Grouped Trials ====================
    ws4 = wb.create_sheet('Grouped Trials')
    ws4.append(trial_headers)
    
    # 按 studyDuration 升序 → setSize 升序分组
    sorted_trials = sorted(all_trials, key=lambda t: (t.get('studyDuration', 0), t.get('setSize', 0)))
    for i, t in enumerate(sorted_trials, 1):
        row = [
            t.get('subjectID', subject_id),
            EXPERIMENT_VERSION,
            EXPERIMENT_NAME,
            sub_group,
            t.get('block', ''),
            t.get('trial', ''),
            i,
            t.get('setSize', ''),
            t.get('studyDuration', ''),
            f"{t.get('studyDuration', '')}ms-SS{t.get('setSize', '')}",
            t.get('accuracy', ''),
            t.get('rt', ''),
            t.get('responseRGB', ''),
            t.get('correctRGB', ''),
            t.get('responseColorIdx', ''),
            t.get('correctColorIdx', ''),
            t.get('trialStatus', 'valid' if not t.get('timeout') else 'timeout'),
            0 if t.get('timeout') else 1,
            t.get('studyDurationActual', ''),
            t.get('changeItem', ''),
            t.get('origColorIdx', ''),
            t.get('probeColorIdx', '')
        ]
        ws4.append(row)
    
    wb.save(file_name)
    print(f"[ PsychoPy ] 数据已导出: {file_name}")
    return file_name

# ============================================================
#  虚拟光标绘制
# ============================================================
def draw_virtual_cursor(win, x, y):
    """绘制虚拟光标：白色圆圈 + 黑色边框"""
    cursor_radius = 9
    # 外圈（白色）
    outer = visual.Circle(win, radius=cursor_radius + 2, fillColor='white', 
                         lineColor='white', lineWidth=3.5, pos=(x, y))
    # 内圈（黑色边框）
    inner = visual.Circle(win, radius=cursor_radius, fillColor='white',
                         lineColor='black', lineWidth=1.5, pos=(x, y))
    outer.draw()
    inner.draw()

# ============================================================
#  主实验流程
# ============================================================
def run_experiment():
    global rng
    
    # 创建窗口 - 全屏，浅灰色背景
    win = visual.Window(
        fullscr=True,
        screen=0,
        winType='pyglet',
        allowGUI=False,
        color=(128/255, 128/255, 128/255),  # rgb(128,128,128) 浅灰色
        colorSpace='rgb',
        units='pix'
    )
    
    # 被试信息
    subject_info = get_subject_info(win)
    sub_group = subject_info['subGroup']
    debug_mode = subject_info['debugMode']
    subject_id = subject_info['subjectID']
    seed = subject_info['rndSeed']
    
    rng = SeededRandom(seed)
    
    # 获取屏幕分辨率
    screen_width, screen_height = win.size
    
    # 刺激参数
    cx, cy = 0, 0  # PsychoPy 全屏模式下原点在屏幕中央
    
    # 创建注视点（黑色圆点，直径 16px = 半径 8px）
    fixation = visual.Circle(win, radius=FIXATION_RADIUS_PX, fillColor='black',
                            lineColor='black', pos=(0, 0))
    
    # 指导语 - 带示意图
    instructions_title = visual.TextStim(win, text='变化定位任务 (CLT)', color='black', height=32, pos=(0, 250))
    
    instructions_text1 = visual.TextStim(win, text='实验流程：', color='black', height=24, pos=(0, 180))
    instructions_text2 = visual.TextStim(win, 
        text='1. 记忆阵列：记住彩色方块的颜色和位置（250ms）', color='black', height=22, pos=(0, 140))
    instructions_text3 = visual.TextStim(win, 
        text='2. 保留间隔：保持注视中央圆点（1000ms）', color='black', height=22, pos=(0, 100))
    instructions_text4 = visual.TextStim(win, 
        text='3. 测试阵列：找出颜色发生变化的方块并点击', color='black', height=22, pos=(0, 60))
    instructions_text5 = visual.TextStim(win, 
        text='注意：请尽快而准确地反应，4秒内未反应记为超时', color='black', height=22, pos=(0, 0))
    instructions_text6 = visual.TextStim(win, 
        text='按空格键开始练习', color='black', height=28, pos=(0, -150))
    
    # 绘制指导语
    instructions_title.draw()
    instructions_text1.draw()
    instructions_text2.draw()
    instructions_text3.draw()
    instructions_text4.draw()
    instructions_text5.draw()
    instructions_text6.draw()
    win.flip()
    event.waitKeys(keyList=['space'])
    
    # 练习结束语
    practice_end_text = visual.TextStim(win, text='', color='black', height=32)
    
    # 休息语
    break_text = visual.TextStim(win, text='休息时间\n\n请放松一下', color='black', height=32)
    break_countdown = visual.TextStim(win, text='', color='red', height=48, pos=(0, -100))
    break_wait = visual.TextStim(win, text='按空格键继续', color='black', height=28, pos=(0, -150))
    
    # 结束语
    finish_text = visual.TextStim(win, text='实验结束\n\n感谢您的参与！', color='black', height=36)
    
    # ==================== 练习阶段 ====================
    print(f"[{EXPERIMENT_NAME}] 开始练习阶段...")
    practice_trials = []
    for ss in SET_SIZES:
        for _ in range(N_PRACTICE_PER_SET):
            practice_trials.append({'setSize': ss, 'studyDuration': STUDY_DURATIONS[0], 'change': 1})
    rng.shuffle(practice_trials)
    
    practice_data = []
    for trial in practice_trials:
        result = run_trial(win, fixation, trial, is_practice=True)
        practice_data.append(result)
    
    # 练习结束统计
    acc = sum(1 for t in practice_data if t.get('accuracy') == 1) / len(practice_data) * 100
    acc_text = f"练习结束\n\n正确率：{acc:.0f}%\n\n按空格键开始正式实验"
    practice_end_text.text = acc_text
    practice_end_text.draw()
    win.flip()
    event.waitKeys(keyList=['space'])
    
    # ==================== 正式实验 ====================
    print(f"[{EXPERIMENT_NAME}] 开始正式实验...")
    
    # 生成正式试次
    formal_trials = []
    for ss in SET_SIZES:
        for _ in range(N_PER_SET_SIZE):
            formal_trials.append({'setSize': ss, 'studyDuration': STUDY_DURATIONS[0], 'change': 1})
    rng.shuffle(formal_trials)
    
    # 分成 blocks
    blocks = [formal_trials[i:i + TRIALS_PER_BLOCK] for i in range(0, len(formal_trials), TRIALS_PER_BLOCK)]
    
    all_data = []
    current_block = 1
    
    for block_idx, block_trials in enumerate(blocks, 1):
        print(f"[{EXPERIMENT_NAME}] Block {block_idx}/{len(blocks)}")
        
        for trial_idx, trial in enumerate(block_trials, 1):
            result = run_trial(win, fixation, trial, is_practice=False, block=block_idx, trial_in_block=trial_idx)
            all_data.append(result)
        
        # Block 间休息（最后一个 block 结束后不休息）
        if block_idx < len(blocks):
            # 强制休息倒计时
            break_text.draw()
            win.flip()
            
            countdown = BREAK_MIN_SECONDS
            while countdown > 0:
                break_countdown.text = f"{countdown}"
                break_text.draw()
                break_countdown.draw()
                win.flip()
                core.wait(1)
                countdown -= 1
            
            # 等待被试按空格
            break_wait.draw()
            win.flip()
            event.waitKeys(keyList=['space'])
    
    # ==================== 结束 ====================
    finish_text.draw()
    win.flip()
    core.wait(2)
    
    # ==================== 导出数据 ====================
    # 补充被试信息
    for t in all_data:
        t['subjectID'] = subject_id
        t['subGroup'] = sub_group
        t['experimentVersion'] = EXPERIMENT_VERSION
        t['experimentType'] = EXPERIMENT_NAME
    
    low_acc = compute_low_acc_flag(all_data)
    file_prefix = 'F' if low_acc['triggered'] and not debug_mode else ''
    export_to_excel(all_data, subject_info, file_prefix)
    
    win.close()
    core.quit()

def run_trial(win, fixation, trial, is_practice=False, block=1, trial_in_block=1):
    """运行单个试次，返回试次数据"""
    set_size = trial['setSize']
    study_duration = trial['studyDuration']
    
    # 生成视觉计划
    plan = build_probe_plan(set_size, rng)
    
    # 生成位置
    screen_width, screen_height = win.size
    locs = get_stim_locs(set_size, screen_width / 2, screen_height / 2, 
                        (screen_width, screen_height), rng)
    plan['locs'] = locs
    
    # 构建试次数据
    trial_data = {
        'block': block,
        'trial': trial_in_block,
        'setSize': set_size,
        'studyDuration': study_duration,
        'changeItem': plan['changeIdx'],
        'origColorIdx': plan['origColor'],
        'probeColorIdx': plan['newColor'],
        'correctColorIdx': plan['newColor'],
        'itemLocs': json.dumps([[round(l['x']), round(l['y'])] for l in locs]),
        'itemColors': json.dumps(plan['memColors']),
        'labelMap': json.dumps(plan['labelMap']),
        'correctResponse': plan['correctLabel']
    }
    
    # ITI（注视点）
    fixation.draw()
    win.flip()
    core.wait(ITI_MS / 1000.0)
    
    # 学习阵列
    study_squares = []
    for i in range(set_size):
        color = COLORS_RGB[plan['memColors'][i]]
        x, y = locs[i]['x'], locs[i]['y']
        square = visual.Rect(win, width=STIM_SIZE_PX, height=STIM_SIZE_PX,
                            fillColor=color, color=color, pos=(x, y))
        study_squares.append(square)
    
    fixation.draw()
    for sq in study_squares:
        sq.draw()
    win.flip()
    core.wait(study_duration / 1000.0)
    
    # 保留间隔（注视点）
    fixation.draw()
    win.flip()
    core.wait(RETENTION_INTERVAL_MS / 1000.0)
    
    # 测试阵列
    test_squares = []
    for i in range(set_size):
        color = COLORS_RGB[plan['probeColors'][i]]
        x, y = locs[i]['x'], locs[i]['y']
        square = visual.Rect(win, width=STIM_SIZE_PX, height=STIM_SIZE_PX,
                            fillColor=color, color=color, pos=(x, y))
        test_squares.append(square)
    
    # 重置虚拟光标到屏幕中央
    virtual_cursor_x = 0
    virtual_cursor_y = 0
    
    # 绘制测试阵列 + 虚拟光标
    fixation.draw()
    for sq in test_squares:
        sq.draw()
    draw_virtual_cursor(win, virtual_cursor_x, virtual_cursor_y)
    win.flip()
    
    # 记录反应开始时间
    response_start = core.getTime()
    mouse = event.Mouse(visible=False, win=win)
    mouse.clickReset()
    
    # 等待反应（最多 4000ms）
    response_recorded = False
    rt = None
    accuracy = 0
    response_color_idx = None
    timeout = False
    
    response_deadline = response_start + RESPONSE_LIMIT_MS / 1000.0
    
    while core.getTime() < response_deadline:
        # 检查鼠标点击
        if mouse.getPressed()[0]:  # 左键
            click_pos = mouse.getPos()
            rt = core.getTime() - response_start
            
            # 判断点击了哪个方块
            clicked_idx = -1
            for i, sq in enumerate(test_squares):
                x, y = sq.pos
                half = STIM_SIZE_PX / 2
                if (x - half <= click_pos[0] <= x + half and
                    y - half <= click_pos[1] <= y + half):
                    clicked_idx = i
                    break
            
            if clicked_idx >= 0:
                response_recorded = True
                response_color_idx = plan['probeColors'][clicked_idx]
                # 判断是否正确：点击的是变化方块
                accuracy = 1 if clicked_idx == plan['changeIdx'] else 0
            break
        
        # 检查是否按了 escape
        if 'escape' in event.getKeys():
            core.quit()
        
        core.wait(0.001)
    
    if not response_recorded:
        timeout = True
        rt = None
        accuracy = 0
        response_color_idx = None
    
    trial_data.update({
        'rt': round(rt, 4) if rt is not None else None,
        'accuracy': accuracy,
        'responseColorIdx': response_color_idx,
        'correctColorIdx': plan['newColor'],
        'responseRGB': ','.join(map(str, COLORS_RGB[response_color_idx])) if response_color_idx is not None else None,
        'correctRGB': ','.join(map(str, COLORS_RGB[plan['newColor']])),
        'timeout': timeout,
        'trialStatus': 'timeout' if timeout else 'valid',
        'countedInK': 0 if timeout else 1
    })
    
    return trial_data

if __name__ == '__main__':
    run_experiment()
