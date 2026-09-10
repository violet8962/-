# -*- coding: utf-8 -*-
"""
《Excel数据可视化——从图表到数据大屏》第二章案例复现
使用 Python (matplotlib/pandas/numpy) 复现 30 个基础图表
"""
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

# ---------------- 全局样式 ----------------
def setup_font():
    candidates = ['Microsoft YaHei', 'SimHei', 'SimSun']
    available = {f.name for f in font_manager.fontManager.ttflist}
    for c in candidates:
        if c in available:
            plt.rcParams['font.sans-serif'] = [c]
            break
    plt.rcParams['axes.unicode_minus'] = False
    plt.rcParams['font.size'] = 11

setup_font()

# Excel 默认配色
C = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5',
     '#70AD47', '#264478', '#9E480E', '#636363', '#997300']

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'charts')
os.makedirs(OUT, exist_ok=True)

# ---------------- 模拟数据 ----------------
months = [f'{m}月' for m in range(1, 13)]
rng = np.random.default_rng(2024)
# 三个产品全年销售额（万元）
prod_a = np.array([120, 135, 158, 142, 168, 185, 196, 210, 188, 172, 195, 225])
prod_b = np.array([85, 92, 105, 98, 112, 125, 132, 138, 128, 118, 130, 145])
prod_c = np.array([60, 68, 72, 75, 82, 88, 92, 96, 90, 84, 95, 105])

quarters = ['一季度', '二季度', '三季度', '四季度']
qa, qb, qc = [413, 495, 594, 592], [282, 335, 398, 393], [200, 245, 278, 284]

# 饼图：各区域市场份额
regions = ['华东', '华南', '华北', '西南', '华中', '其他']
shares = [32, 22, 18, 12, 10, 6]

# 散点/气泡：广告投入 vs 销售额
n = 12
ad = np.linspace(20, 200, n) + rng.normal(0, 6, n)
sales = 80 + 1.6 * ad + rng.normal(0, 25, n)
profits = np.abs(40 + 0.5 * ad + rng.normal(0, 20, n))  # 气泡大小：利润

# 股价（开盘-盘高-盘低-收盘），5个交易日
stock = pd.DataFrame({
    '日期': ['9/2', '9/3', '9/4', '9/5', '9/6'],
    '开盘': [10.20, 10.55, 10.80, 10.65, 11.00],
    '最高': [10.80, 11.05, 11.20, 11.10, 11.45],
    '最低': [10.05, 10.40, 10.60, 10.45, 10.85],
    '收盘': [10.60, 10.90, 10.70, 11.05, 11.30],
})

# 雷达图：两款产品六维评分
dims = ['动力', '油耗', '空间', '舒适', '安全', '价格']
car_x = [9, 7, 8, 8, 9, 7]
car_y = [7, 9, 7, 9, 8, 9]

# 瀑布图：年度利润桥（万元）
wf_items = ['期初利润', '产品A增收', '产品B增收', '成本上升', '新市场增收', '税费支出', '期末利润']
wf_vals = [500, 180, 120, -90, 150, -70, 790]

data_sheets = {}  # 用于导出 数据源.xlsx

def save_fig(fig, idx, name):
    path = os.path.join(OUT, f'{idx:02d}_{name}.png')
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f'[{idx:02d}] {name}')
    return path

def add_sheet(idx, name, df):
    data_sheets[f'{idx:02d}_{name}'] = df

# ============ 1 簇状柱形图 ============
fig, ax = plt.subplots(figsize=(9, 5))
x = np.arange(len(quarters)); w = 0.25
ax.bar(x - w, qa, w, label='产品A', color=C[0])
ax.bar(x, qb, w, label='产品B', color=C[1])
ax.bar(x + w, qc, w, label='产品C', color=C[2])
ax.set_title('各季度产品销售额对比（簇状柱形图）'); ax.set_ylabel('销售额（万元）')
ax.set_xticks(x); ax.set_xticklabels(quarters); ax.legend()
add_sheet(1, '簇状柱形图', pd.DataFrame({'季度': quarters, '产品A': qa, '产品B': qb, '产品C': qc}))
save_fig(fig, 1, '簇状柱形图')

# ============ 2 堆积柱形图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.bar(quarters, qa, label='产品A', color=C[0])
ax.bar(quarters, qb, bottom=qa, label='产品B', color=C[1])
ax.bar(quarters, qc, bottom=np.array(qa) + np.array(qb), label='产品C', color=C[2])
ax.set_title('各季度销售额构成（堆积柱形图）'); ax.set_ylabel('销售额（万元）'); ax.legend()
add_sheet(2, '堆积柱形图', pd.DataFrame({'季度': quarters, '产品A': qa, '产品B': qb, '产品C': qc}))
save_fig(fig, 2, '堆积柱形图')

# ============ 3 百分比堆积柱形图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ta = np.array(qa, float); tb = np.array(qb, float); tc = np.array(qc, float)
tot = ta + tb + tc
ax.bar(quarters, ta / tot * 100, label='产品A', color=C[0])
ax.bar(quarters, tb / tot * 100, bottom=ta / tot * 100, label='产品B', color=C[1])
ax.bar(quarters, tc / tot * 100, bottom=(ta + tb) / tot * 100, label='产品C', color=C[2])
ax.set_title('各季度销售额占比（百分比堆积柱形图）'); ax.set_ylabel('占比（%）'); ax.set_ylim(0, 100); ax.legend(loc='upper right', bbox_to_anchor=(1, 1))
add_sheet(3, '百分比堆积柱形图', pd.DataFrame({'季度': quarters, '产品A占比%': np.round(ta/tot*100,1), '产品B占比%': np.round(tb/tot*100,1), '产品C占比%': np.round(tc/tot*100,1)}))
save_fig(fig, 3, '百分比堆积柱形图')

# ============ 4 三维簇状柱形图 ============
fig = plt.figure(figsize=(9, 6))
ax = fig.add_subplot(111, projection='3d')
xs = np.arange(4)
for i, (vals, color, lab) in enumerate([(qa, C[0], '产品A'), (qb, C[1], '产品B'), (qc, C[2], '产品C')]):
    ax.bar3d(xs, i, 0, 0.5, 0.5, vals, color=color, alpha=0.85)
ax.set_xticks(xs + 0.25); ax.set_xticklabels(quarters)
ax.set_yticks([0.25, 1.25, 2.25]); ax.set_yticklabels(['产品A', '产品B', '产品C'])
ax.set_zlabel('销售额（万元）'); ax.set_title('各季度产品销售额（三维簇状柱形图）')
add_sheet(4, '三维柱形图', pd.DataFrame({'季度': quarters, '产品A': qa, '产品B': qb, '产品C': qc}))
save_fig(fig, 4, '三维簇状柱形图')

# ============ 5 簇状条形图 ============
fig, ax = plt.subplots(figsize=(9, 5))
y = np.arange(len(quarters)); h = 0.25
ax.barh(y - h, qa, h, label='产品A', color=C[0])
ax.barh(y, qb, h, label='产品B', color=C[1])
ax.barh(y + h, qc, h, label='产品C', color=C[2])
ax.set_title('各季度产品销售额对比（簇状条形图）'); ax.set_xlabel('销售额（万元）')
ax.set_yticks(y); ax.set_yticklabels(quarters); ax.legend()
add_sheet(5, '簇状条形图', pd.DataFrame({'季度': quarters, '产品A': qa, '产品B': qb, '产品C': qc}))
save_fig(fig, 5, '簇状条形图')

# ============ 6 堆积条形图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.barh(quarters, qa, label='产品A', color=C[0])
ax.barh(quarters, qb, left=qa, label='产品B', color=C[1])
ax.barh(quarters, qc, left=np.array(qa) + np.array(qb), label='产品C', color=C[2])
ax.set_title('各季度销售额构成（堆积条形图）'); ax.set_xlabel('销售额（万元）'); ax.legend()
add_sheet(6, '堆积条形图', pd.DataFrame({'季度': quarters, '产品A': qa, '产品B': qb, '产品C': qc}))
save_fig(fig, 6, '堆积条形图')

# ============ 7 百分比堆积条形图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.barh(quarters, ta / tot * 100, label='产品A', color=C[0])
ax.barh(quarters, tb / tot * 100, left=ta / tot * 100, label='产品B', color=C[1])
ax.barh(quarters, tc / tot * 100, left=(ta + tb) / tot * 100, label='产品C', color=C[2])
ax.set_title('各季度销售额占比（百分比堆积条形图）'); ax.set_xlabel('占比（%）'); ax.set_xlim(0, 100); ax.legend()
add_sheet(7, '百分比堆积条形图', pd.DataFrame({'季度': quarters, '产品A占比%': np.round(ta/tot*100,1), '产品B占比%': np.round(tb/tot*100,1), '产品C占比%': np.round(tc/tot*100,1)}))
save_fig(fig, 7, '百分比堆积条形图')

# ============ 8 折线图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.plot(months, prod_a, color=C[0], lw=2, label='产品A')
ax.plot(months, prod_b, color=C[1], lw=2, label='产品B')
ax.plot(months, prod_c, color=C[2], lw=2, label='产品C')
ax.set_title('全年月度销售趋势（折线图）'); ax.set_ylabel('销售额（万元）'); ax.legend(); ax.grid(alpha=0.3)
add_sheet(8, '折线图', pd.DataFrame({'月份': months, '产品A': prod_a, '产品B': prod_b, '产品C': prod_c}))
save_fig(fig, 8, '折线图')

# ============ 9 堆积折线图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.stackplot(months, prod_a, prod_b, prod_c, labels=['产品A', '产品B', '产品C'], colors=C[:3], alpha=0.85)
ax.set_title('全年累计销售额（堆积折线图）'); ax.set_ylabel('累计销售额（万元）'); ax.legend(loc='upper left'); ax.grid(alpha=0.3)
add_sheet(9, '堆积折线图', pd.DataFrame({'月份': months, '产品A': prod_a, '产品B': prod_b, '产品C': prod_c}))
save_fig(fig, 9, '堆积折线图')

# ============ 10 百分比堆积折线图 ============
fig, ax = plt.subplots(figsize=(9, 5))
pa = prod_a / (prod_a + prod_b + prod_c) * 100
pb = prod_b / (prod_a + prod_b + prod_c) * 100
pc = prod_c / (prod_a + prod_b + prod_c) * 100
ax.stackplot(months, pa, pb, pc, labels=['产品A', '产品B', '产品C'], colors=C[:3], alpha=0.85)
ax.set_title('全年销售额占比趋势（百分比堆积折线图）'); ax.set_ylabel('占比（%）'); ax.set_ylim(0, 100); ax.legend(loc='upper left'); ax.grid(alpha=0.3)
add_sheet(10, '百分比堆积折线图', pd.DataFrame({'月份': months, '产品A占比%': np.round(pa,1), '产品B占比%': np.round(pb,1), '产品C占比%': np.round(pc,1)}))
save_fig(fig, 10, '百分比堆积折线图')

# ============ 11 带数据标记的折线图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.plot(months, prod_a, color=C[0], lw=2, marker='o', markersize=6, label='产品A')
ax.plot(months, prod_b, color=C[1], lw=2, marker='s', markersize=6, label='产品B')
ax.plot(months, prod_c, color=C[3], lw=2, marker='^', markersize=7, label='产品C')
ax.set_title('全年月度销售趋势（带数据标记的折线图）'); ax.set_ylabel('销售额（万元）'); ax.legend(); ax.grid(alpha=0.3)
add_sheet(11, '数据标记折线图', pd.DataFrame({'月份': months, '产品A': prod_a, '产品B': prod_b, '产品C': prod_c}))
save_fig(fig, 11, '带数据标记的折线图')

# ============ 12 饼图 ============
fig, ax = plt.subplots(figsize=(8, 6))
ax.pie(shares, labels=regions, autopct='%1.1f%%', startangle=90, colors=C[:6],
       wedgeprops={'edgecolor': 'white'}, textprops={'fontsize': 11})
ax.set_title('各区域市场份额（饼图）')
add_sheet(12, '饼图', pd.DataFrame({'区域': regions, '占比%': shares}))
save_fig(fig, 12, '饼图')

# ============ 13 分离型饼图 ============
fig, ax = plt.subplots(figsize=(8, 6))
explode = [0, 0, 0, 0, 0, 0.12]
ax.pie(shares, labels=regions, autopct='%1.1f%%', startangle=90, colors=C[:6],
       explode=explode, wedgeprops={'edgecolor': 'white'}, textprops={'fontsize': 11})
ax.set_title('各区域市场份额（分离型饼图）')
add_sheet(13, '分离型饼图', pd.DataFrame({'区域': regions, '占比%': shares, '分离': ['否']*5 + ['是']}))
save_fig(fig, 13, '分离型饼图')

# ============ 14 复合饼图 ============
fig = plt.figure(figsize=(10, 5.5))
ax1 = fig.add_axes([0.02, 0.05, 0.55, 0.9])
main_vals = shares[:5] + [sum(shares[5:])]
main_labels = regions[:5] + ['其他']
wedges, _ = ax1.pie(main_vals, labels=main_labels, startangle=90, colors=C[:5] + [C[9]],
                    wedgeprops={'edgecolor': 'white'}, textprops={'fontsize': 11})
ax1.set_title('各区域市场份额（复合饼图）')
# 分离最后一个扇区的视觉强调
wedges[-1].set_edgecolor('white')
ax2 = fig.add_axes([0.62, 0.25, 0.33, 0.5])
sub_vals = [4, 2]  # 其他=西北4%、东北2%（合计6%）
ax2.pie(sub_vals, labels=['西北', '东北'], autopct='%1.0f%%', startangle=90,
        colors=[C[9], C[10] if len(C) > 10 else '#7030A0'],
        wedgeprops={'edgecolor': 'white'}, textprops={'fontsize': 10})
ax2.set_title('其他区域明细', fontsize=11)
add_sheet(14, '复合饼图', pd.DataFrame({'区域': regions[:5] + ['西北', '东北'], '占比%': shares[:5] + [4, 2]}))
save_fig(fig, 14, '复合饼图')

# ============ 15 复合条饼图 ============
fig = plt.figure(figsize=(10, 5.5))
ax1 = fig.add_axes([0.02, 0.05, 0.55, 0.9])
ax1.pie(main_vals, labels=main_labels, startangle=90, colors=C[:5] + [C[9]],
        wedgeprops={'edgecolor': 'white'}, textprops={'fontsize': 11})
ax1.set_title('各区域市场份额（复合条饼图）')
ax2 = fig.add_axes([0.62, 0.3, 0.33, 0.4])
ax2.bar(['西北', '东北'], sub_vals, color=[C[9], '#7030A0'], width=0.5)
ax2.set_title('其他区域构成', fontsize=11); ax2.set_ylabel('占比（%）')
for i, v in enumerate(sub_vals):
    ax2.text(i, v + 0.1, f'{v}%', ha='center', fontsize=10)
add_sheet(15, '复合条饼图', pd.DataFrame({'区域': regions[:5] + ['西北', '东北'], '占比%': shares[:5] + sub_vals}))
save_fig(fig, 15, '复合条饼图')

# ============ 16 圆环图 ============
fig, ax = plt.subplots(figsize=(8, 6))
ax.pie(shares, labels=regions, autopct='%1.1f%%', startangle=90, colors=C[:6],
       wedgeprops={'width': 0.4, 'edgecolor': 'white'}, textprops={'fontsize': 11})
ax.set_title('各区域市场份额（圆环图）')
add_sheet(16, '圆环图', pd.DataFrame({'区域': regions, '占比%': shares}))
save_fig(fig, 16, '圆环图')

# ============ 17 分离型圆环图 ============
fig, ax = plt.subplots(figsize=(8, 6))
ax.pie(shares, labels=regions, autopct='%1.1f%%', startangle=90, colors=C[:6],
       explode=[0, 0, 0, 0, 0, 0.12],
       wedgeprops={'width': 0.4, 'edgecolor': 'white'}, textprops={'fontsize': 11})
ax.set_title('各区域市场份额（分离型圆环图）')
add_sheet(17, '分离型圆环图', pd.DataFrame({'区域': regions, '占比%': shares}))
save_fig(fig, 17, '分离型圆环图')

# ============ 18 面积图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.fill_between(months, prod_a, color=C[0], alpha=0.4, label='产品A')
ax.plot(months, prod_a, color=C[0], lw=1.5)
ax.set_title('产品A全年销售走势（面积图）'); ax.set_ylabel('销售额（万元）'); ax.legend(); ax.grid(alpha=0.3)
add_sheet(18, '面积图', pd.DataFrame({'月份': months, '产品A': prod_a}))
save_fig(fig, 18, '面积图')

# ============ 19 堆积面积图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.stackplot(months, prod_a, prod_b, prod_c, labels=['产品A', '产品B', '产品C'], colors=C[:3], alpha=0.7)
ax.set_title('全年销售额构成（堆积面积图）'); ax.set_ylabel('销售额（万元）'); ax.legend(loc='upper left'); ax.grid(alpha=0.3)
add_sheet(19, '堆积面积图', pd.DataFrame({'月份': months, '产品A': prod_a, '产品B': prod_b, '产品C': prod_c}))
save_fig(fig, 19, '堆积面积图')

# ============ 20 百分比堆积面积图 ============
fig, ax = plt.subplots(figsize=(9, 5))
ax.stackplot(months, pa, pb, pc, labels=['产品A', '产品B', '产品C'], colors=C[:3], alpha=0.7)
ax.set_title('全年销售额占比（百分比堆积面积图）'); ax.set_ylabel('占比（%）'); ax.set_ylim(0, 100); ax.legend(loc='upper left'); ax.grid(alpha=0.3)
add_sheet(20, '百分比堆积面积图', pd.DataFrame({'月份': months, '产品A占比%': np.round(pa,1), '产品B占比%': np.round(pb,1), '产品C占比%': np.round(pc,1)}))
save_fig(fig, 20, '百分比堆积面积图')

# ============ 21 散点图 ============
fig, ax = plt.subplots(figsize=(8, 5.5))
ax.scatter(ad, sales, color=C[0], s=60, alpha=0.75, edgecolors='white')
ax.set_xlabel('广告投入（万元）'); ax.set_ylabel('销售额（万元）')
ax.set_title('广告投入与销售额关系（散点图）'); ax.grid(alpha=0.3)
add_sheet(21, '散点图', pd.DataFrame({'广告投入': np.round(ad, 1), '销售额': np.round(sales, 1)}))
save_fig(fig, 21, '散点图')

# ============ 22 带直线和数据标记的散点图 ============
fig, ax = plt.subplots(figsize=(8, 5.5))
idx = np.argsort(ad)
ax.plot(ad[idx], sales[idx], color=C[1], lw=1.5, marker='o', markersize=6)
ax.set_xlabel('广告投入（万元）'); ax.set_ylabel('销售额（万元）')
ax.set_title('广告投入与销售额关系（带直线和数据标记的散点图）'); ax.grid(alpha=0.3)
add_sheet(22, '带直线散点图', pd.DataFrame({'广告投入': np.round(np.sort(ad), 1), '销售额': np.round(sales[idx], 1)}))
save_fig(fig, 22, '带直线和数据标记的散点图')

# ============ 23 带平滑线的散点图 ============
from scipy.interpolate import make_interp_spline
fig, ax = plt.subplots(figsize=(8, 5.5))
xs_s = np.sort(ad)
ys_s = sales[np.argsort(ad)]
spl = make_interp_spline(xs_s, ys_s, k=3)
x_smooth = np.linspace(xs_s.min(), xs_s.max(), 200)
ax.scatter(xs_s, ys_s, color=C[6], s=50, zorder=3)
ax.plot(x_smooth, spl(x_smooth), color=C[6], lw=2)
ax.set_xlabel('广告投入（万元）'); ax.set_ylabel('销售额（万元）')
ax.set_title('广告投入与销售额关系（带平滑线的散点图）'); ax.grid(alpha=0.3)
save_fig(fig, 23, '带平滑线的散点图')

# ============ 24 气泡图 ============
fig, ax = plt.subplots(figsize=(8.5, 5.5))
sc = ax.scatter(ad, sales, s=profits * 8, color=C[5], alpha=0.55, edgecolors='white')
ax.set_xlabel('广告投入（万元）'); ax.set_ylabel('销售额（万元）')
ax.set_title('广告投入、销售额与利润关系（气泡图，气泡大小=利润）'); ax.grid(alpha=0.3)
add_sheet(24, '气泡图', pd.DataFrame({'广告投入': np.round(ad, 1), '销售额': np.round(sales, 1), '利润(气泡)': np.round(profits, 1)}))
save_fig(fig, 24, '气泡图')

# ============ 25 股价图（开盘-盘高-盘低-收盘） ============
fig, ax = plt.subplots(figsize=(9, 5.5))
for i, row in stock.iterrows():
    color = '#C00000' if row['收盘'] >= row['开盘'] else '#008000'  # 红涨绿跌
    ax.plot([i, i], [row['最低'], row['最高']], color=color, lw=1.8)          # 高低线
    ax.plot([i - 0.15, i], [row['开盘'], row['开盘']], color=color, lw=2.5)    # 开盘左横线
    ax.plot([i, i + 0.15], [row['收盘'], row['收盘']], color=color, lw=2.5)    # 收盘右横线
ax.set_xticks(range(len(stock))); ax.set_xticklabels(stock['日期'])
ax.set_ylabel('股价（元）'); ax.set_title('某股票5日行情（开盘-盘高-盘低-收盘股价图）'); ax.grid(alpha=0.3)
add_sheet(25, '股价图', stock)
save_fig(fig, 25, '股价图')

# ============ 26 雷达图 ============
fig, ax = plt.subplots(figsize=(7, 7), subplot_kw=dict(polar=True))
ang = np.linspace(0, 2 * np.pi, len(dims), endpoint=False).tolist()
ang += ang[:1]
x_v = car_x + car_x[:1]; y_v = car_y + car_y[:1]
ax.plot(ang, x_v, color=C[0], lw=2, label='车型X')
ax.plot(ang, y_v, color=C[1], lw=2, label='车型Y')
ax.set_xticks(ang[:-1]); ax.set_xticklabels(dims); ax.set_ylim(0, 10)
ax.set_title('两款车型六维性能对比（雷达图）', pad=20); ax.legend(loc='upper right', bbox_to_anchor=(1.2, 1.1))
add_sheet(26, '雷达图', pd.DataFrame({'维度': dims, '车型X': car_x, '车型Y': car_y}))
save_fig(fig, 26, '雷达图')

# ============ 27 带数据标记的雷达图 ============
fig, ax = plt.subplots(figsize=(7, 7), subplot_kw=dict(polar=True))
ax.plot(ang, x_v, color=C[0], lw=2, marker='o', markersize=6, label='车型X')
ax.plot(ang, y_v, color=C[1], lw=2, marker='s', markersize=6, label='车型Y')
ax.set_xticks(ang[:-1]); ax.set_xticklabels(dims); ax.set_ylim(0, 10)
ax.set_title('两款车型六维性能对比（带数据标记的雷达图）', pad=20); ax.legend(loc='upper right', bbox_to_anchor=(1.2, 1.1))
save_fig(fig, 27, '带数据标记的雷达图')

# ============ 28 填充雷达图 ============
fig, ax = plt.subplots(figsize=(7, 7), subplot_kw=dict(polar=True))
ax.fill(ang, x_v, color=C[0], alpha=0.3, label='车型X')
ax.fill(ang, y_v, color=C[1], alpha=0.3, label='车型Y')
ax.plot(ang, x_v, color=C[0], lw=1.5)
ax.plot(ang, y_v, color=C[1], lw=1.5)
ax.set_xticks(ang[:-1]); ax.set_xticklabels(dims); ax.set_ylim(0, 10)
ax.set_title('两款车型六维性能对比（填充雷达图）', pad=20); ax.legend(loc='upper right', bbox_to_anchor=(1.2, 1.1))
save_fig(fig, 28, '填充雷达图')

# ============ 29 组合图（柱形+折线，次坐标轴） ============
fig, ax1 = plt.subplots(figsize=(9, 5))
total_monthly = prod_a + prod_b + prod_c
growth = np.r_[np.nan, np.diff(total_monthly) / total_monthly[:-1] * 100]
ax1.bar(months, total_monthly, color=C[4], alpha=0.85, label='月销售额')
ax1.set_ylabel('销售额（万元）')
ax2 = ax1.twinx()
ax2.plot(months, growth, color=C[7] if len(C) > 7 else '#C00000', lw=2, marker='o', markersize=5, label='环比增长率')
ax2.set_ylabel('环比增长率（%）')
ax1.set_title('月度销售额与环比增长率（柱形+折线组合图）'); ax1.grid(alpha=0.3)
lines1, labels1 = ax1.get_legend_handles_labels(); lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
add_sheet(29, '组合图', pd.DataFrame({'月份': months, '销售额': total_monthly, '环比增长率%': np.round(growth, 1)}))
save_fig(fig, 29, '组合图')

# ============ 30 瀑布图 ============
fig, ax = plt.subplots(figsize=(10, 5.5))
cum = 0; bottoms = []; heights = []; colors_w = []
labels_w = []
for i, (lab, v) in enumerate(zip(wf_items, wf_vals)):
    if i == 0 or i == len(wf_vals) - 1:
        bottoms.append(0); heights.append(v); colors_w.append(C[0]); cum = v
    else:
        if v >= 0:
            bottoms.append(cum); heights.append(v); colors_w.append(C[5])
            cum += v
        else:
            cum += v
            bottoms.append(cum); heights.append(-v); colors_w.append(C[1])
    labels_w.append(lab)
xb = np.arange(len(wf_items))
ax.bar(xb, heights, bottom=bottoms, color=colors_w, edgecolor='white', width=0.6)
# 连接线
cum2 = 0
for i, v in enumerate(wf_vals):
    if i == 0:
        cum2 = v
    elif i == len(wf_vals) - 1:
        pass
    else:
        prev = cum2
        cum2 += v
        ax.plot([i - 0.3, i + 0.3], [prev, prev], color='gray', lw=0.8, ls='--')
# 数据标签
for i, (b, h, v) in enumerate(zip(bottoms, heights, wf_vals)):
    ax.text(i, b + h + 12, f'{v:+d}' if 0 < i < len(wf_vals) - 1 else f'{v}', ha='center', fontsize=9)
ax.set_xticks(xb); ax.set_xticklabels(labels_w, fontsize=10)
ax.set_ylabel('利润（万元）'); ax.set_title('年度利润构成分析（瀑布图）')
add_sheet(30, '瀑布图', pd.DataFrame({'项目': wf_items, '金额(万元)': wf_vals}))
save_fig(fig, 30, '瀑布图')

# ---------------- 导出数据源 Excel ----------------
xlsx_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '数据源.xlsx')
with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
    for sheet, df in data_sheets.items():
        # sheet 名最长31字符
        df.to_excel(writer, sheet_name=sheet[:31], index=False)
print(f'\n数据源已保存: {xlsx_path}')
print(f'共生成 {len(os.listdir(OUT))} 张图表，位于: {OUT}')
