"""
达人投放数据看板生成器 v1.0
读取Excel数据模板，生成交互式HTML分析报告
用法：python generate_dashboard.py [数据文件路径] [输出文件路径]
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd
import json
from datetime import datetime
import os

def load_data(filepath):
    """读取投放数据，支持模板格式（投放明细表sheet）和按月分sheet格式"""
    xls = pd.ExcelFile(filepath, engine='openpyxl')
    sheets = xls.sheet_names

    all_dfs = []
    if '投放明细表' in sheets:
        # 模板格式：单sheet
        df = pd.read_excel(xls, sheet_name='投放明细表', header=1)
        all_dfs.append(df)
    else:
        # 按月分sheet格式：1月、2月...
        for sheet in sheets:
            df = pd.read_excel(xls, sheet_name=sheet)
            if len(df) > 0:
                all_dfs.append(df)

    df = pd.concat(all_dfs, ignore_index=True)

    # ===== 列名统一映射 =====
    col_map = {}
    # 日期列
    for c in df.columns:
        cl = str(c).strip()
        if cl in ['日期', '投放日期']:
            col_map[c] = '投放日期'
        elif cl in ['发布链接', '内容链接', '笔记链接']:
            col_map[c] = '内容链接'
        elif cl in ['款式', '款式名称']:
            col_map[c] = '款式名称'
        elif cl in ['款号', '款式编号']:
            col_map[c] = '款式编号'
        elif cl in ['笔记', '达人姓名', '达人ID'] and '达人姓名' not in col_map.values():
            # 「笔记」列在原始数据里存的是达人/笔记标识，这里当作达人姓名字段
            if cl == '笔记':
                col_map[c] = '达人姓名'
            elif cl == '达人姓名':
                col_map[c] = '达人姓名'
        elif cl in ['投放费用', '达人合作费']:
            col_map[c] = '达人合作费'
        elif cl in ['成本+运费', '样品/运费成本']:
            col_map[c] = '样品/运费成本'
        elif cl in ['点赞', '点赞数']:
            col_map[c] = '点赞数'
        elif cl in ['收藏', '收藏数']:
            col_map[c] = '收藏数'
        elif cl in ['评论', '评论数']:
            col_map[c] = '评论数'
        elif cl in ['平台', '内容类型', '点击量', '所属周']:
            col_map[c] = cl

    df = df.rename(columns=col_map)

    # 如果没有平台投放费列，补一列0
    if '平台投放费' not in df.columns:
        df['平台投放费'] = 0

    # 如果没有所属周列，后面再算
    # 去掉全空行（按投放日期）
    df = df.dropna(subset=['投放日期'])
    # 转换日期
    df['投放日期'] = pd.to_datetime(df['投放日期'])

    # 确保数值列是数字
    for col in ['达人合作费', '平台投放费', '样品/运费成本', '点击量', '点赞数', '收藏数', '评论数']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 如果没有达人姓名列，补空字符串
    if '达人姓名' not in df.columns:
        df['达人姓名'] = ''

    # 计算总费用、总互动
    df['总费用'] = df['达人合作费'].fillna(0) + df['平台投放费'].fillna(0)
    df['总互动'] = df['点赞数'].fillna(0) + df['收藏数'].fillna(0) + df['评论数'].fillna(0)
    # 月份
    df['月份'] = df['投放日期'].dt.month
    df['月份标签'] = df['投放日期'].dt.strftime('%Y-%m')
    # 周
    df['周标签'] = df['投放日期'].dt.strftime('%Y-W%W')
    # 所属周
    df['所属周'] = df['投放日期'].dt.strftime('第%W周')

    # ===== 款号合并（同一款的不同写法统一）=====
    style_alias_map = {
        '001A1': '0001A1',
        '000A1': '0001A1',
    }
    df['款式编号'] = df['款式编号'].astype(str).replace(style_alias_map)

    # ===== 款式系列划分 =====
    utri_styles = ['495', '0002A1', '0001A1', '0003A2', '493AX', '492']
    df['款式系列'] = df['款式编号'].apply(
        lambda s: 'UTRI系列' if s in utri_styles else '中台系列'
    )

    return df


def build_chart_data(df):
    """构建所有图表数据"""
    data = {}

    # ===== 概览 =====
    total_cost = df['总费用'].sum()
    total_clicks = df['点击量'].sum()
    total_inter = df['总互动'].sum()
    total_likes = df['点赞数'].sum()
    total_favs = df['收藏数'].sum()
    total_comments = df['评论数'].sum()
    avg_rate = total_inter / total_clicks if total_clicks > 0 else 0
    avg_cpi = total_cost / total_inter if total_inter > 0 else 0
    post_count = len(df)
    style_count = df['款式编号'].nunique()

    data['overview'] = {
        'totalCost': round(total_cost, 2),
        'totalClicks': int(total_clicks),
        'totalInteractions': int(total_inter),
        'totalLikes': int(total_likes),
        'totalFavorites': int(total_favs),
        'totalComments': int(total_comments),
        'avgRate': round(avg_rate * 100, 2),
        'avgCPI': round(avg_cpi, 4),
        'postCount': post_count,
        'styleCount': style_count,
        'dateRange': f"{df['投放日期'].min().strftime('%Y-%m-%d')} ~ {df['投放日期'].max().strftime('%Y-%m-%d')}",
        'updateDate': datetime.now().strftime('%Y-%m-%d'),
    }

    # ===== 月度趋势 =====
    monthly = df.groupby('月份标签').agg(
        费用=('总费用', 'sum'),
        互动量=('总互动', 'sum'),
        点击量=('点击量', 'sum'),
        笔记数=('投放日期', 'count'),
    ).reset_index()
    monthly['互动率'] = monthly.apply(lambda r: r['互动量'] / r['点击量'] if r['点击量'] > 0 else 0, axis=1)
    monthly['单互动成本'] = monthly.apply(lambda r: r['费用'] / r['互动量'] if r['互动量'] > 0 else 0, axis=1)

    data['monthly'] = {
        'labels': monthly['月份标签'].tolist(),
        'cost': [round(v, 1) for v in monthly['费用'].tolist()],
        'interactions': [int(v) for v in monthly['互动量'].tolist()],
        'clicks': [int(v) for v in monthly['点击量'].tolist()],
        'posts': [int(v) for v in monthly['笔记数'].tolist()],
        'rate': [round(v * 100, 2) for v in monthly['互动率'].tolist()],
        'cpi': [round(v, 4) for v in monthly['单互动成本'].tolist()],
    }

    # ===== 周度趋势 =====
    weekly = df.groupby('周标签').agg(
        费用=('总费用', 'sum'),
        互动量=('总互动', 'sum'),
        点击量=('点击量', 'sum'),
        笔记数=('投放日期', 'count'),
    ).reset_index().sort_values('周标签')
    weekly['互动率'] = weekly.apply(lambda r: r['互动量'] / r['点击量'] if r['点击量'] > 0 else 0, axis=1)
    weekly['单互动成本'] = weekly.apply(lambda r: r['费用'] / r['互动量'] if r['互动量'] > 0 else 0, axis=1)
    # 生成友好的周标签
    weekly_friendly = []
    for _, row in weekly.iterrows():
        # 从周标签中提取周数，格式如 2026-W05
        parts = row['周标签'].split('-W')
        if len(parts) == 2:
            week_num = int(parts[1])
            weekly_friendly.append('第' + str(week_num) + '周')
        else:
            weekly_friendly.append(row['周标签'])
    weekly['周标签友好'] = weekly_friendly

    data['weekly'] = {
        'labels': weekly['周标签友好'].tolist(),
        'rawLabels': weekly['周标签'].tolist(),
        'cost': [round(v, 1) for v in weekly['费用'].tolist()],
        'interactions': [int(v) for v in weekly['互动量'].tolist()],
        'clicks': [int(v) for v in weekly['点击量'].tolist()],
        'posts': [int(v) for v in weekly['笔记数'].tolist()],
        'rate': [round(v * 100, 2) for v in weekly['互动率'].tolist()],
        'cpi': [round(v, 4) for v in weekly['单互动成本'].tolist()],
    }

    # ===== 内容类型分析 =====
    ctype = df.groupby('内容类型').agg(
        笔记数=('投放日期', 'count'),
        费用=('总费用', 'sum'),
        互动量=('总互动', 'sum'),
        点击量=('点击量', 'sum'),
    ).reset_index()
    ctype['互动率'] = ctype.apply(lambda r: r['互动量'] / r['点击量'] if r['点击量'] > 0 else 0, axis=1)
    ctype['单互动成本'] = ctype.apply(lambda r: r['费用'] / r['互动量'] if r['互动量'] > 0 else 999, axis=1)
    ctype = ctype.sort_values('费用', ascending=False)

    # TOP20
    top_ctype = ctype.head(20)
    data['ctype'] = {
        'names': top_ctype['内容类型'].tolist(),
        'cost': [round(v, 1) for v in top_ctype['费用'].tolist()],
        'posts': [int(v) for v in top_ctype['笔记数'].tolist()],
        'interactions': [int(v) for v in top_ctype['互动量'].tolist()],
        'rate': [round(v * 100, 2) for v in top_ctype['互动率'].tolist()],
        'cpi': [round(v, 4) for v in top_ctype['单互动成本'].tolist()],
        'totalTypes': len(ctype),
    }

    # 内容大类（自动分类）
    celeb_keywords = [
        '刘耀文', '陈奕恒', '赵露思', '马嘉祺', '张桂源', '赵一博', '刘轩丞',
        '穆祉丞', '杨紫', '侯明昊', '左奇函', '邓紫棋', '陈浚铭', '鞠婧祎',
        '杨博文', '王一珩', '江衡', '高桥大翔', '李沛恩', '星邱', '迪丽热巴',
        '官俊臣', '魏笑', '张子墨', '王橹杰', '虞书欣', '左航', '程相',
        '朱映宸', '肖战', '汪苏泷', '李煜东', '秦彻', '林俊杰', '陈立农',
        '苏新皓', '文严文', '李沛恩江衡'
    ]

    def classify(ct):
        if ct in celeb_keywords:
            return '明星达人'
        if '穿搭' in str(ct) or ct in ['学生', '健身', '通勤', '情侣', '女高', '上班', '地铁', '群像']:
            return '穿搭类'
        if '口播' in str(ct):
            return '口播类'
        if any(k in str(ct) for k in ['合集', '好物', '开箱', '分享', 'vlog', '思路', '姿势', '应援色', '野餐']):
            return '种草合集类'
        if ct == '跳舞':
            return '舞蹈类'
        return '其他'

    df['内容大类'] = df['内容类型'].apply(classify)
    cat_group = df.groupby('内容大类').agg(
        笔记数=('投放日期', 'count'),
        费用=('总费用', 'sum'),
        互动量=('总互动', 'sum'),
        点击量=('点击量', 'sum'),
    ).reset_index().sort_values('费用', ascending=False)
    cat_group['互动率'] = cat_group.apply(lambda r: r['互动量'] / r['点击量'] if r['点击量'] > 0 else 0, axis=1)
    cat_group['单互动成本'] = cat_group.apply(lambda r: r['费用'] / r['互动量'] if r['互动量'] > 0 else 0, axis=1)

    data['category'] = {
        'names': cat_group['内容大类'].tolist(),
        'cost': [round(v, 1) for v in cat_group['费用'].tolist()],
        'posts': [int(v) for v in cat_group['笔记数'].tolist()],
        'interactions': [int(v) for v in cat_group['互动量'].tolist()],
        'rate': [round(v * 100, 2) for v in cat_group['互动率'].tolist()],
        'cpi': [round(v, 4) for v in cat_group['单互动成本'].tolist()],
    }

    # 月度内容大类堆叠
    month_cat = df.groupby(['月份标签', '内容大类'])['总费用'].sum().reset_index()
    cat_names = cat_group['内容大类'].tolist()
    monthly_cat = {}
    for cat in cat_names:
        sub = month_cat[month_cat['内容大类'] == cat]
        vals = []
        for m in monthly['月份标签']:
            v = sub[sub['月份标签'] == m]['总费用'].values
            vals.append(round(float(v[0]), 1) if len(v) > 0 else 0)
        monthly_cat[cat] = vals
    data['monthlyCategory'] = monthly_cat

    # ===== 款式系列分析 =====
    series = df.groupby('款式系列').agg(
        笔记数=('投放日期', 'count'),
        费用=('总费用', 'sum'),
        互动量=('总互动', 'sum'),
        点击量=('点击量', 'sum'),
        款式数=('款式编号', 'nunique'),
    ).reset_index().sort_values('费用', ascending=False)
    series['互动率'] = series.apply(lambda r: r['互动量'] / r['点击量'] if r['点击量'] > 0 else 0, axis=1)
    series['单互动成本'] = series.apply(lambda r: r['费用'] / r['互动量'] if r['互动量'] > 0 else 0, axis=1)

    data['series'] = {
        'names': series['款式系列'].tolist(),
        'cost': [round(v, 1) for v in series['费用'].tolist()],
        'posts': [int(v) for v in series['笔记数'].tolist()],
        'interactions': [int(v) for v in series['互动量'].tolist()],
        'clicks': [int(v) for v in series['点击量'].tolist()],
        'styles': [int(v) for v in series['款式数'].tolist()],
        'rate': [round(v * 100, 2) for v in series['互动率'].tolist()],
        'cpi': [round(v, 4) for v in series['单互动成本'].tolist()],
    }

    # 月度系列趋势
    month_series = df.groupby(['月份标签', '款式系列'])['总费用'].sum().reset_index()
    monthly_series_data = {}
    for s_name in series['款式系列'].tolist():
        sub = month_series[month_series['款式系列'] == s_name]
        vals = []
        for m in monthly['月份标签']:
            v = sub[sub['月份标签'] == m]['总费用'].values
            vals.append(round(float(v[0]), 1) if len(v) > 0 else 0)
        monthly_series_data[s_name] = vals
    data['monthlySeries'] = monthly_series_data

    # ===== 款式分析 =====
    style = df.groupby('款式编号').agg(
        笔记数=('投放日期', 'count'),
        费用=('总费用', 'sum'),
        互动量=('总互动', 'sum'),
        点击量=('点击量', 'sum'),
    ).reset_index()
    style['互动率'] = style.apply(lambda r: r['互动量'] / r['点击量'] if r['点击量'] > 0 else 0, axis=1)
    style['单互动成本'] = style.apply(lambda r: r['费用'] / r['互动量'] if r['互动量'] > 0 else 999, axis=1)
    style = style.sort_values('费用', ascending=False)

    data['style'] = {
        'names': style['款式编号'].tolist(),
        'cost': [round(v, 1) for v in style['费用'].tolist()],
        'posts': [int(v) for v in style['笔记数'].tolist()],
        'interactions': [int(v) for v in style['互动量'].tolist()],
        'rate': [round(v * 100, 2) for v in style['互动率'].tolist()],
        'cpi': [round(v, 4) for v in style['单互动成本'].tolist()],
    }

    # ===== 平台分析 =====
    plat = df.groupby('平台').agg(
        笔记数=('投放日期', 'count'),
        费用=('总费用', 'sum'),
        互动量=('总互动', 'sum'),
        点击量=('点击量', 'sum'),
    ).reset_index().sort_values('费用', ascending=False)
    plat['互动率'] = plat.apply(lambda r: r['互动量'] / r['点击量'] if r['点击量'] > 0 else 0, axis=1)
    plat['单互动成本'] = plat.apply(lambda r: r['费用'] / r['互动量'] if r['互动量'] > 0 else 0, axis=1)

    data['platform'] = {
        'names': plat['平台'].tolist(),
        'cost': [round(v, 1) for v in plat['费用'].tolist()],
        'posts': [int(v) for v in plat['笔记数'].tolist()],
        'interactions': [int(v) for v in plat['互动量'].tolist()],
        'rate': [round(v * 100, 2) for v in plat['互动率'].tolist()],
        'cpi': [round(v, 4) for v in plat['单互动成本'].tolist()],
    }

    # ===== 明细数据（用于表格筛选）=====
    detail_cols = ['投放日期', '所属周', '平台', '达人姓名', '款式系列', '款式编号', '内容类型',
                   '内容链接', '达人合作费', '平台投放费', '样品/运费成本', '点击量', '点赞数', '收藏数', '评论数']
    detail = df[detail_cols].copy()
    detail['投放日期'] = detail['投放日期'].dt.strftime('%Y-%m-%d')
    detail['总费用'] = (df['达人合作费'].fillna(0) + df['平台投放费'].fillna(0)).round(2)
    detail['总互动'] = df['总互动'].astype(int)
    detail['互动率'] = (df['总互动'] / df['点击量'].where(df['点击量'] > 0)).round(4) * 100
    detail['单互动成本'] = (df['总费用'] / df['总互动'].where(df['总互动'] > 0)).round(4)
    data['detail'] = detail.values.tolist()
    data['detailColumns'] = list(detail.columns)

    return data


def generate_html(data, output_path):
    """生成HTML看板"""
    html = HTML_TEMPLATE.replace('/*__DATA__*/', 'const ALL_DATA = ' + json.dumps(data, ensure_ascii=False) + ';')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'✅ 看板已生成：{output_path}')


# ============================================================
# HTML 模板
# ============================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>达人投放数据看板</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif;background:#f7f8fa;color:#1a1d21;padding:24px}
.container{max-width:1400px;margin:0 auto}
h1{font-size:24px;font-weight:700;margin-bottom:4px}
.subtitle{color:#6b7280;margin-bottom:6px;font-size:13px}
.update-date{color:#9ca3af;margin-bottom:20px;font-size:12px}
.update-date::before{content:'🗓️ 更新日期：'}

/* 筛选栏 */
.filter-bar{background:#fff;border-radius:12px;padding:16px 20px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.04);border:1px solid #eef0f3;display:flex;gap:16px;align-items:center;flex-wrap:wrap}
.filter-label{font-size:13px;color:#6b7280;font-weight:500}
.filter-select{padding:6px 12px;border:1px solid #d1d5db;border-radius:6px;font-size:13px;background:#fff;cursor:pointer;font-family:inherit}
.filter-select:focus{outline:none;border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.1)}

/* KPI */
.kpi-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:20px}
.kpi-card{background:#fff;border-radius:12px;padding:16px 18px;box-shadow:0 1px 3px rgba(0,0,0,.04);border:1px solid #eef0f3}
.kpi-label{font-size:12px;color:#6b7280;margin-bottom:6px}
.kpi-value{font-size:22px;font-weight:700;color:#1a1d21}
.kpi-unit{font-size:12px;color:#6b7280;margin-left:3px;font-weight:400}
.kpi-sub{font-size:11px;color:#9ca3af;margin-top:4px}
.kpi-mom{font-size:11px;margin-top:6px;font-weight:500}
.kpi-mom.pos{color:#10b981}
.kpi-mom.neg{color:#ef4444}
.kpi-mom.flat{color:#9ca3af}

/* 图表区 */
.chart-section{background:#fff;border-radius:12px;padding:20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.04);border:1px solid #eef0f3}
.section-title{font-size:16px;font-weight:600;margin-bottom:16px;padding-left:10px;border-left:3px solid #6366f1}
.chart-row{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.chart-row-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px}
.chart-box{position:relative;height:300px}
.chart-box-tall{position:relative;height:380px}
.chart-box-wide{position:relative;height:320px}

/* 表格 */
.table-section{background:#fff;border-radius:12px;padding:20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.04);border:1px solid #eef0f3}
.table-toolbar{display:flex;gap:12px;margin-bottom:12px;align-items:center;flex-wrap:wrap}
.table-toolbar input{padding:6px 12px;border:1px solid #d1d5db;border-radius:6px;font-size:13px;font-family:inherit;width:200px}
.table-toolbar input:focus{outline:none;border-color:#6366f1}
table{width:100%;border-collapse:collapse;font-size:12px}
th,td{padding:8px 10px;text-align:right;border-bottom:1px solid #f3f4f6;white-space:nowrap}
th:first-child,td:first-child{text-align:left}
th{background:#f9fafb;font-weight:600;color:#4b5563;font-size:11px;position:sticky;top:0;cursor:pointer;user-select:none}
th:hover{background:#f3f4f6}
tr:hover{background:#f9fafb}
.table-wrap{max-height:500px;overflow:auto;border:1px solid #f3f4f6;border-radius:8px}
.table-count{font-size:12px;color:#6b7280;margin-left:auto}

/* 爆款榜单 */
.top-toolbar{display:flex;gap:12px;margin-bottom:14px;align-items:center;flex-wrap:wrap}
.top-desc{font-size:12px;color:#6b7280;margin-bottom:12px;padding:8px 12px;background:#f9fafb;border-radius:6px;border-left:3px solid #6366f1;line-height:1.6}
.top-desc-icon{margin-right:4px}
.top-desc-threshold{color:#9ca3af;margin-left:6px}
.top-table-wrap{overflow-x:auto;border:1px solid #f3f4f6;border-radius:8px}
.top-table-wrap table{width:100%;border-collapse:collapse;font-size:12px}
.top-table-wrap th,.top-table-wrap td{padding:10px 12px;text-align:center;border-bottom:1px solid #f3f4f6;white-space:nowrap}
.top-table-wrap th:first-child,.top-table-wrap td:first-child{text-align:left;width:30px}
.top-table-wrap th{background:#f9fafb;font-weight:600;color:#4b5563;font-size:11px}
.top-table-wrap tr:hover{background:#f9fafb}
.top-rank{display:inline-block;width:22px;height:22px;line-height:22px;text-align:center;border-radius:50%;font-weight:700;font-size:11px}
.top-rank-1{background:#fef3c7;color:#92400e}
.top-rank-2{background:#f3f4f6;color:#4b5563}
.top-rank-3{background:#fde68a;color:#92400e}
.top-rank-n{background:#f3f4f6;color:#6b7280}
.top-link{display:inline-block;padding:3px 10px;background:#eef2ff;color:#4f46e5;border-radius:4px;font-size:11px;text-decoration:none;white-space:nowrap}
.top-link:hover{background:#e0e7ff}

/* 颜色标签 */
.cat-dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px;vertical-align:middle}

@media(max-width:1100px){
  .kpi-grid{grid-template-columns:repeat(3,1fr)}
  .chart-row,.chart-row-3{grid-template-columns:1fr}
}
@media(max-width:600px){
  .kpi-grid{grid-template-columns:repeat(2,1fr)}
  body{padding:12px}
}
</style>
</head>
<body>
<div class="container">
  <h1>📊 达人投放数据看板</h1>
  <p class="subtitle" id="dateRange">-</p>
  <p class="update-date" id="updateDate">-</p>

  <!-- 筛选栏 -->
  <div class="filter-bar">
    <span class="filter-label">时间范围：</span>
    <select class="filter-select" id="timeFilter" onchange="applyFilter()">
      <option value="all">全部数据</option>
      <option value="month">本月</option>
      <option value="lastMonth">上月</option>
      <option value="week7">近7天</option>
      <option value="week30">近30天</option>
    </select>
    <span class="filter-label">月份：</span>
    <select class="filter-select" id="monthFilter" onchange="applyFilter()">
      <option value="all">全部月份</option>
    </select>
    <span class="filter-label">趋势维度：</span>
    <select class="filter-select" id="trendDim" onchange="renderTrend()">
      <option value="monthly">按月</option>
      <option value="weekly">按周</option>
    </select>
    <span class="filter-label">平台：</span>
    <select class="filter-select" id="platFilter" onchange="applyFilter()">
      <option value="all">全部平台</option>
    </select>
    <span class="filter-label">系列：</span>
    <select class="filter-select" id="seriesFilter" onchange="applyFilter()">
      <option value="all">全部系列</option>
    </select>
    <span class="filter-label">款式：</span>
    <select class="filter-select" id="styleFilter" onchange="applyFilter()">
      <option value="all">全部款式</option>
    </select>
  </div>

  <!-- KPI -->
  <div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-label">总投放费用</div><div class="kpi-value" id="kpiCost">-</div><div class="kpi-sub" id="kpiPosts">共 - 条笔记</div><div class="kpi-mom" id="momCost">-</div></div>
    <div class="kpi-card"><div class="kpi-label">总点击量</div><div class="kpi-value" id="kpiClicks">-</div><div class="kpi-sub">单次点击成本 <span id="kpiCPC">-</span></div><div class="kpi-mom" id="momClicks">-</div></div>
    <div class="kpi-card"><div class="kpi-label">总互动量</div><div class="kpi-value" id="kpiInter">-</div><div class="kpi-sub">点赞+收藏+评论</div><div class="kpi-mom" id="momInter">-</div></div>
    <div class="kpi-card"><div class="kpi-label">平均互动率</div><div class="kpi-value" id="kpiRate">-</div><div class="kpi-sub">行业参考 3-8%</div><div class="kpi-mom" id="momRate">-</div></div>
    <div class="kpi-card"><div class="kpi-label">单互动成本</div><div class="kpi-value" id="kpiCPI">-</div><div class="kpi-sub">越低越好</div><div class="kpi-mom" id="momCPI">-</div></div>
    <div class="kpi-card"><div class="kpi-label">覆盖款式</div><div class="kpi-value" id="kpiStyles">-</div><div class="kpi-sub">个款式</div><div class="kpi-mom" id="momStyles">-</div></div>
  </div>

  <!-- 趋势分析 -->
  <div class="chart-section">
    <div class="section-title">📈 趋势分析</div>
    <div class="chart-row">
      <div class="chart-box"><canvas id="trendCost"></canvas></div>
      <div class="chart-box"><canvas id="trendInter"></canvas></div>
    </div>
  </div>

  <!-- 内容大类 -->
  <div class="chart-section">
    <div class="section-title">🎯 内容大类分析</div>
    <div class="chart-row">
      <div class="chart-box"><canvas id="catCost"></canvas></div>
      <div class="chart-box"><canvas id="catEff"></canvas></div>
    </div>
    <div style="margin-top:16px">
      <div class="chart-box-wide"><canvas id="catStacked"></canvas></div>
    </div>
  </div>

  <!-- 内容类型明细 -->
  <div class="chart-section">
    <div class="section-title">🎬 内容类型分析 TOP20</div>
    <div class="chart-row">
      <div class="chart-box-tall"><canvas id="ctypeCost"></canvas></div>
      <div class="chart-box-tall"><canvas id="ctypeScatter"></canvas></div>
    </div>
  </div>

  <!-- 款式系列分析 -->
  <div class="chart-section">
    <div class="section-title">👟 款式系列分析</div>
    <div class="chart-row">
      <div class="chart-box"><canvas id="seriesCost"></canvas></div>
      <div class="chart-box"><canvas id="seriesEff"></canvas></div>
    </div>
    <div style="margin-top:16px">
      <div class="chart-box-wide"><canvas id="seriesTrend"></canvas></div>
    </div>
  </div>

  <!-- 款式分析 -->
  <div class="chart-section">
    <div class="section-title">🔢 款式明细分析</div>
    <div class="chart-row">
      <div class="chart-box-tall"><canvas id="styleCost"></canvas></div>
      <div class="chart-box-tall"><canvas id="styleScatter"></canvas></div>
    </div>
  </div>

  <!-- 平台分析 -->
  <div class="chart-section">
    <div class="section-title">📱 平台分析</div>
    <div class="chart-row">
      <div class="chart-box"><canvas id="platCost"></canvas></div>
      <div class="chart-box"><canvas id="platEff"></canvas></div>
    </div>
  </div>

  <!-- 爆款笔记榜 -->
  <div class="chart-section">
    <div class="section-title">🏆 月度爆款笔记榜 TOP10</div>
    <div class="top-desc">
      <span class="top-desc-icon">💡</span>
      <span>综合评分规则：分平台百分位排名，权重 = 点击量30% + 互动率30% + 互动量25% + 单互动成本15%</span>
      <span class="top-desc-threshold">｜入围门槛：抖音≥1000点击/50互动，小红书≥300点击/20互动</span>
    </div>
    <div class="top-toolbar">
      <span class="filter-label">月份：</span>
      <select class="filter-select" id="topMonthFilter" onchange="renderTopNotes()">
        <option value="all">全部月份</option>
      </select>
      <span class="filter-label">排序：</span>
      <select class="filter-select" id="topSortFilter" onchange="renderTopNotes()">
        <option value="score">综合分</option>
        <option value="rate">按互动率</option>
        <option value="inter">按互动量</option>
        <option value="cpi">按单互动成本</option>
      </select>
    </div>
    <div class="top-table-wrap">
      <table id="topTable">
        <thead><tr id="topTableHead"></tr></thead>
        <tbody id="topTableBody"></tbody>
      </table>
    </div>
  </div>

  <!-- 明细表格 -->
  <div class="table-section">
    <div class="section-title">📋 投放明细</div>
    <div class="table-toolbar">
      <input type="text" id="tableSearch" placeholder="🔍 搜索达人/款式/平台/类型..." oninput="renderTable()">
      <span class="table-count" id="tableCount">- 条记录</span>
    </div>
    <div class="table-wrap">
      <table id="detailTable">
        <thead><tr id="tableHead"></tr></thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
/*__DATA__*/

// ============ 颜色 ============
const CAT_COLORS = ['#6366f1','#10b981','#f59e0b','#ef4444','#8b5cf6','#ec4899','#06b6d4','#64748b'];
const PLAT_COLORS = ['#3b82f6','#ef4444','#10b981','#f59e0b','#8b5cf6','#ec4899','#06b6d4','#84cc16'];
Chart.defaults.font.family='-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif';
Chart.defaults.font.size=11;
Chart.defaults.color='#6b7280';

// ============ 状态 ============
let filteredData = null;
let sortCol = -1, sortDir = 1;
let currentPage = 0;
const PAGE_SIZE = 100;

// ============ 筛选 ============
function applyFilter() {
  const timeVal = document.getElementById('timeFilter').value;
  const monthVal = document.getElementById('monthFilter').value;
  const platVal = document.getElementById('platFilter').value;
  const seriesVal = document.getElementById('seriesFilter').value;
  const styleVal = document.getElementById('styleFilter').value;

  const all = ALL_DATA.detail.map(r => {
    const obj = {};
    ALL_DATA.detailColumns.forEach((c, i) => obj[c] = r[i]);
    return obj;
  });

  let filtered = all;

  // 平台筛选
  if (platVal !== 'all') {
    filtered = filtered.filter(r => r['平台'] === platVal);
  }
  // 系列筛选
  if (seriesVal !== 'all') {
    filtered = filtered.filter(r => r['款式系列'] === seriesVal);
  }
  // 款式筛选
  if (styleVal !== 'all') {
    filtered = filtered.filter(r => r['款式编号'] === styleVal);
  }
  // 时间筛选
  const now = new Date();
  if (timeVal === 'week7') {
    const d7 = new Date(now.getTime() - 7*86400000);
    filtered = filtered.filter(r => new Date(r['投放日期']) >= d7);
  } else if (timeVal === 'week30') {
    const d30 = new Date(now.getTime() - 30*86400000);
    filtered = filtered.filter(r => new Date(r['投放日期']) >= d30);
  } else if (timeVal === 'month') {
    const ym = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0');
    filtered = filtered.filter(r => r['投放日期'].startsWith(ym));
  } else if (timeVal === 'lastMonth') {
    const d = new Date(now.getFullYear(), now.getMonth()-1, 1);
    const ym = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
    filtered = filtered.filter(r => r['投放日期'].startsWith(ym));
  }
  // 月份筛选（具体某月）
  if (monthVal !== 'all') {
    filtered = filtered.filter(r => r['投放日期'].startsWith(monthVal));
  }

  // 重新计算过滤后的汇总数据
  filteredData = filtered;
  const summary = computeSummary(filtered);
  renderKPIs(summary);
  renderTrend();
  renderCategoryCharts(filtered);
  renderCtypeCharts(filtered);
  renderSeriesCharts(filtered);
  renderStyleCharts(filtered);
  renderPlatformCharts(filtered);
  renderTopNotes();
  renderTable();
}

function computeSummary(rows) {
  const cost = rows.reduce((s,r) => s + (r['总费用']||0), 0);
  const clicks = rows.reduce((s,r) => s + (r['点击量']||0), 0);
  const inter = rows.reduce((s,r) => s + (r['总互动']||0), 0);
  const styles = new Set(rows.map(r => r['款式编号'])).size;
  return {
    cost, clicks, inter,
    rate: clicks > 0 ? inter / clicks * 100 : 0,
    cpi: inter > 0 ? cost / inter : 0,
    cpc: clicks > 0 ? cost / clicks : 0,
    posts: rows.length,
    styles,
  };
}

// 计算环比：将筛选后的数据按时间对半分，后半段 vs 前半段
function computeMoM(rows) {
  if (rows.length < 2) return null;
  const sorted = rows.slice().sort((a, b) => new Date(a['投放日期']) - new Date(b['投放日期']));
  const mid = Math.floor(sorted.length / 2);
  const firstHalf = sorted.slice(0, mid);
  const secondHalf = sorted.slice(mid);
  if (firstHalf.length === 0 || secondHalf.length === 0) return null;
  const cur = computeSummary(secondHalf);
  const prev = computeSummary(firstHalf);
  const mom = (key, lowerIsBetter = false) => {
    const cv = cur[key], pv = prev[key];
    if (pv === 0 || pv === undefined) return { pct: null, dir: 'flat', arrow: '—' };
    const raw = (cv - pv) / pv * 100;
    const abs = Math.abs(raw).toFixed(1);
    let dir, arrow;
    if (Math.abs(raw) < 0.5) {
      dir = 'flat'; arrow = '—';
    } else if (lowerIsBetter) {
      // 越低越好的指标（费用、成本）：降了=pos（绿），涨了=neg（红）
      if (raw > 0) { dir = 'neg'; arrow = '↑'; }
      else { dir = 'pos'; arrow = '↓'; }
    } else {
      // 越高越好的指标（互动、点击、率）：涨了=pos（绿），降了=neg（红）
      if (raw > 0) { dir = 'pos'; arrow = '↑'; }
      else { dir = 'neg'; arrow = '↓'; }
    }
    return { pct: abs, dir, arrow };
  };
  return {
    cost: mom('cost', true),     // 费用：降了好
    clicks: mom('clicks'),       // 点击：涨了好
    inter: mom('inter'),         // 互动：涨了好
    rate: mom('rate'),           // 互动率：涨了好
    cpi: mom('cpi', true),       // 单互动成本：降了好
    styles: mom('styles'),       // 款式数：涨了好
  };
}

function formatMoM(momData, key) {
  if (!momData || !momData[key]) return '环比 -';
  const m = momData[key];
  if (m.pct === null) return '环比 -';
  return `环比 ${m.arrow} ${m.pct}%`;
}

function renderKPIs(s) {
  const mom = computeMoM(filteredData);
  document.getElementById('kpiCost').innerHTML = '¥' + Math.round(s.cost).toLocaleString() + '<span class="kpi-unit">元</span>';
  document.getElementById('kpiPosts').textContent = '共 ' + s.posts + ' 条笔记';
  document.getElementById('momCost').textContent = formatMoM(mom, 'cost');
  document.getElementById('momCost').className = 'kpi-mom ' + (mom?.cost?.dir || 'flat');

  document.getElementById('kpiClicks').innerHTML = (s.clicks/10000).toFixed(1) + '<span class="kpi-unit">万</span>';
  document.getElementById('kpiCPC').textContent = '¥' + s.cpc.toFixed(4);
  document.getElementById('momClicks').textContent = formatMoM(mom, 'clicks');
  document.getElementById('momClicks').className = 'kpi-mom ' + (mom?.clicks?.dir || 'flat');

  document.getElementById('kpiInter').innerHTML = (s.inter/10000).toFixed(1) + '<span class="kpi-unit">万</span>';
  document.getElementById('momInter').textContent = formatMoM(mom, 'inter');
  document.getElementById('momInter').className = 'kpi-mom ' + (mom?.inter?.dir || 'flat');

  document.getElementById('kpiRate').innerHTML = s.rate.toFixed(2) + '<span class="kpi-unit">%</span>';
  document.getElementById('momRate').textContent = formatMoM(mom, 'rate');
  document.getElementById('momRate').className = 'kpi-mom ' + (mom?.rate?.dir || 'flat');

  document.getElementById('kpiCPI').innerHTML = '¥' + s.cpi.toFixed(4);
  document.getElementById('momCPI').textContent = formatMoM(mom, 'cpi');
  document.getElementById('momCPI').className = 'kpi-mom ' + (mom?.cpi?.dir || 'flat');

  document.getElementById('kpiStyles').textContent = s.styles;
  document.getElementById('momStyles').textContent = formatMoM(mom, 'styles');
  document.getElementById('momStyles').className = 'kpi-mom ' + (mom?.styles?.dir || 'flat');

  document.getElementById('dateRange').textContent = ALL_DATA.overview.dateRange + ' · 共 ' + ALL_DATA.overview.postCount + ' 条笔记 · ' + ALL_DATA.overview.styleCount + ' 个款式';
  document.getElementById('updateDate').textContent = ALL_DATA.overview.updateDate;
}

// ============ 趋势图 ============
let trendCostChart, trendInterChart;
function getMonthLabel(dateStr) {
  return dateStr.slice(0, 7);
}
function getWeekLabel(dateStr) {
  const d = new Date(dateStr);
  // ISO周数
  const target = new Date(d.valueOf());
  const dayNr = (d.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = target.valueOf();
  target.setMonth(0, 1);
  if (target.getDay() !== 4) {
    target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);
  }
  const weekNum = 1 + Math.ceil((firstThursday - target) / 604800000);
  return '第' + weekNum + '周';
}
function aggregateTrend(rows, dim) {
  const map = {};
  rows.forEach(r => {
    const label = dim === 'monthly' ? getMonthLabel(r['投放日期']) : getWeekLabel(r['投放日期']);
    if (!map[label]) map[label] = { cost: 0, inter: 0, clicks: 0, posts: 0 };
    map[label].cost += r['总费用'] || 0;
    map[label].inter += r['总互动'] || 0;
    map[label].clicks += r['点击量'] || 0;
    map[label].posts += 1;
  });
  let labels = Object.keys(map).sort();
  // 按月标签保持 YYYY-MM 格式用于排序，显示时用中文
  if (dim === 'monthly') {
    labels = labels.sort();
  }
  const costs = labels.map(l => Math.round(map[l].cost * 10) / 10);
  const inters = labels.map(l => Math.round(map[l].inter));
  const rates = labels.map(l => map[l].clicks > 0 ? Math.round(map[l].inter / map[l].clicks * 10000) / 100 : 0);
  const cpis = labels.map(l => map[l].inter > 0 ? Math.round(map[l].cost / map[l].inter * 10000) / 10000 : 0);
  const posts = labels.map(l => map[l].posts);
  const displayLabels = dim === 'monthly'
    ? labels.map(l => l.replace('-', '年') + '月')
    : labels;
  return { labels: displayLabels, costs, inters, rates, cpis, posts };
}
function renderTrend() {
  const dim = document.getElementById('trendDim').value;
  const rows = filteredData || ALL_DATA.detail.map(r => {
    const obj = {};
    ALL_DATA.detailColumns.forEach((c, i) => obj[c] = r[i]);
    return obj;
  });
  const agg = aggregateTrend(rows, dim);

  if (trendCostChart) trendCostChart.destroy();
  if (trendInterChart) trendInterChart.destroy();

  trendCostChart = new Chart(document.getElementById('trendCost'), {
    type: 'bar',
    data: { labels: agg.labels, datasets: [{ label: '投放费用(元)', data: agg.costs, backgroundColor: 'rgba(99,102,241,0.85)', borderRadius: 4, barThickness: 28 }]},
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{display:false}, title:{display:true,text:dim==='monthly'?'月度投放费用':'周度投放费用',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}} },
      scales: { x:{grid:{display:false}}, y:{grid:{color:'#f3f4f6'},ticks:{callback:v=>'¥'+v.toLocaleString()}} }
    }
  });

  trendInterChart = new Chart(document.getElementById('trendInter'), {
    type: 'line',
    data: {
      labels: agg.labels,
      datasets: [
        { label:'总互动量', data: agg.inters, borderColor:'#10b981', backgroundColor:'rgba(16,185,129,0.1)', fill:true, tension:.3, pointRadius:4, yAxisID:'y' },
        { label:'互动率(%)', data: agg.rates, borderColor:'#f59e0b', borderDash:[4,4], fill:false, tension:.3, pointRadius:4, yAxisID:'y1' }
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      interaction:{mode:'index',intersect:false},
      plugins: { title:{display:true,text:dim==='monthly'?'月度互动量 & 互动率':'周度互动量 & 互动率',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'bottom'} },
      scales: { x:{grid:{display:false}}, y:{grid:{color:'#f3f4f6'},ticks:{callback:v=>(v/10000).toFixed(0)+'万'}}, y1:{position:'right',grid:{display:false},ticks:{callback:v=>v+'%'}} }
    }
  });
}

// ============ 内容大类 ============
let catCostChart, catEffChart, catStackedChart;
function renderCategoryCharts(rows) {
  const catMap = {};
  const catKeyMap = {};
  // 用预计算的分类名映射
  rows.forEach(r => {
    const ct = r['内容类型'];
    const cat = ALL_DATA.category.names.includes(ct) ? ct : classifyContent(ct);
    if (!catMap[cat]) catMap[cat] = { cost:0, inter:0, clicks:0, posts:0 };
    catMap[cat].cost += r['总费用']||0;
    catMap[cat].inter += r['总互动']||0;
    catMap[cat].clicks += r['点击量']||0;
    catMap[cat].posts += 1;
  });
  const cats = Object.keys(catMap).sort((a,b) => catMap[b].cost - catMap[a].cost);
  const costs = cats.map(c => Math.round(catMap[c].cost*10)/10);
  const rates = cats.map(c => catMap[c].clicks>0 ? catMap[c].inter/catMap[c].clicks*100 : 0);
  const cpis = cats.map(c => catMap[c].inter>0 ? catMap[c].cost/catMap[c].inter : 0);

  if (catCostChart) catCostChart.destroy();
  catCostChart = new Chart(document.getElementById('catCost'), {
    type:'doughnut',
    data:{ labels:cats, datasets:[{ data:costs, backgroundColor:CAT_COLORS, borderWidth:2, borderColor:'#fff', borderRadius:4 }] },
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ title:{display:true,text:'内容大类费用占比',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'right',labels:{padding:10}} }
    }
  });

  if (catEffChart) catEffChart.destroy();
  catEffChart = new Chart(document.getElementById('catEff'), {
    type:'bar',
    data:{ labels:cats, datasets:[
      { label:'互动率(%)', data:rates.map(v=>Math.round(v*100)/100), backgroundColor:'rgba(16,185,129,0.85)', borderRadius:4, barThickness:24, yAxisID:'y' },
      { label:'单互动成本(元)', data:cpis.map(v=>Math.round(v*10000)/10000), backgroundColor:'rgba(239,68,68,0.85)', borderRadius:4, barThickness:24, yAxisID:'y1' }
    ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ title:{display:true,text:'各大类效率对比',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'bottom'} },
      scales:{ x:{grid:{display:false}}, y:{position:'left',grid:{color:'#f3f4f6'},ticks:{callback:v=>v+'%'}}, y1:{position:'right',grid:{display:false},ticks:{callback:v=>'¥'+v}} }
    }
  });

  // 堆叠图用原始月度数据
  if (catStackedChart) catStackedChart.destroy();
  const rawLabels = ALL_DATA.monthly.labels;
  const displayLabels = rawLabels.map(l => l.replace('-', '年') + '月');
  const datasets = cats.map((cat, i) => {
    const monthlyData = rawLabels.map(label => {
      const monthRows = rows.filter(r => r['投放日期'].startsWith(label.slice(0,7)));
      let sum = 0;
      monthRows.forEach(r => {
        const rc = classifyContent(r['内容类型']);
        if (rc === cat) sum += r['总费用']||0;
      });
      return Math.round(sum*10)/10;
    });
    return { label:cat, data:monthlyData, backgroundColor:CAT_COLORS[i%CAT_COLORS.length], borderRadius:0, barThickness:24 };
  });
  catStackedChart = new Chart(document.getElementById('catStacked'), {
    type:'bar',
    data:{ labels: displayLabels, datasets },
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ title:{display:true,text:'月度费用结构（各大类）',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'bottom'} },
      scales:{ x:{stacked:true,grid:{display:false}}, y:{stacked:true,grid:{color:'#f3f4f6'},ticks:{callback:v=>'¥'+v.toLocaleString()}} }
    }
  });
}

const CELEB_LIST = ['刘耀文','陈奕恒','赵露思','马嘉祺','张桂源','赵一博','刘轩丞','穆祉丞','杨紫','侯明昊','左奇函','邓紫棋','陈浚铭','鞠婧祎','杨博文','王一珩','江衡','高桥大翔','李沛恩','星邱','迪丽热巴','官俊臣','魏笑','张子墨','王橹杰','虞书欣','左航','程相','朱映宸','肖战','汪苏泷','李煜东','秦彻','林俊杰','陈立农','苏新皓','文严文','李沛恩江衡'];
function classifyContent(ct) {
  if (CELEB_LIST.includes(ct)) return '明星达人';
  if (typeof ct === 'string' && ct.includes('穿搭')) return '穿搭类';
  if (['学生','健身','通勤','情侣','女高','上班','地铁','群像'].includes(ct)) return '穿搭类';
  if (typeof ct === 'string' && ct.includes('口播')) return '口播类';
  if (typeof ct === 'string' && ['合集','好物','开箱','分享','vlog','思路','姿势','应援色','野餐'].some(k=>ct.includes(k))) return '种草合集类';
  if (ct === '跳舞') return '舞蹈类';
  return '其他';
}

// ============ 内容类型TOP20 ============
let ctypeCostChart, ctypeScatterChart;
function renderCtypeCharts(rows) {
  const map = {};
  rows.forEach(r => {
    const ct = r['内容类型'] || '未知';
    if (!map[ct]) map[ct] = { cost:0, inter:0, clicks:0, posts:0 };
    map[ct].cost += r['总费用']||0;
    map[ct].inter += r['总互动']||0;
    map[ct].clicks += r['点击量']||0;
    map[ct].posts += 1;
  });
  const arr = Object.entries(map).map(([name,d]) => ({name,...d}))
    .sort((a,b) => b.cost - a.cost).slice(0, 20);
  const names = arr.map(x=>x.name);
  const costs = arr.map(x=>Math.round(x.cost*10)/10);
  const rates = arr.map(x=>x.clicks>0?x.inter/x.clicks*100:0);
  const cpis = arr.map(x=>x.inter>0?x.cost/x.inter:999);
  const maxCost = Math.max(...costs, 1);

  if (ctypeCostChart) ctypeCostChart.destroy();
  ctypeCostChart = new Chart(document.getElementById('ctypeCost'), {
    type:'bar',
    data:{ labels:names, datasets:[{ label:'费用(元)', data:costs, backgroundColor:'rgba(236,72,153,0.85)', borderRadius:3, barThickness:18 }]},
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, title:{display:true,text:'内容类型费用 TOP20',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}} },
      scales:{ x:{grid:{color:'#f3f4f6'},ticks:{callback:v=>'¥'+v.toLocaleString()}}, y:{grid:{display:false}} }
    }
  });

  if (ctypeScatterChart) ctypeScatterChart.destroy();
  // 计算中位数作为四象限分割线
  const sortedCpis = [...cpis].sort((a,b) => a-b);
  const sortedRates = [...rates].sort((a,b) => a-b);
  const midCpi = sortedCpis[Math.floor(sortedCpis.length/2)];
  const midRate = sortedRates[Math.floor(sortedRates.length/2)];

  // 四象限颜色
  // 明星=绿(高互动率+低成本) 问题=橙(高互动率+高成本) 淘汰=红(低互动率+高成本) 现金牛=黄(低互动率+低成本)
  const scatterColors = arr.map((x, i) => {
    const isHighRate = rates[i] >= midRate;
    const isLowCost = cpis[i] <= midCpi;
    if (isHighRate && isLowCost) return { bg: 'rgba(16,185,129,0.6)', border: '#10b981' };      // 明星
    if (isHighRate && !isLowCost) return { bg: 'rgba(245,158,11,0.6)', border: '#f59e0b' };    // 问题
    if (!isHighRate && !isLowCost) return { bg: 'rgba(239,68,68,0.6)', border: '#ef4444' };   // 淘汰
    return { bg: 'rgba(139,92,246,0.5)', border: '#8b5cf6' };                                  // 现金牛
  });

  ctypeScatterChart = new Chart(document.getElementById('ctypeScatter'), {
    type:'scatter',
    data:{ datasets:[{
      label:'内容类型',
      data: arr.map((x,i) => ({
        x: Math.min(cpis[i], 10),
        y: Math.round(rates[i]*100)/100,
        r: 4 + (costs[i]/maxCost)*12,
        name: x.name,
        posts: x.posts,
      })),
      backgroundColor: scatterColors.map(c => c.bg),
      borderColor: scatterColors.map(c => c.border),
      borderWidth: 1.5,
    }]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, title:{display:true,text:'效率矩阵（气泡=费用，颜色=象限）',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}},
        tooltip:{ callbacks:{ label:ctx=>{ const d=ctx.raw; return [d.name,'笔记: '+d.posts+'条','互动率: '+d.y+'%','单互动成本: ¥'+Math.round(d.x*10000)/10000]; } }}
      },
      scales:{
        x:{ title:{display:true,text:'单互动成本(元) ← 越左越优'}, grid:{color:'#f3f4f6'} },
        y:{ title:{display:true,text:'互动率(%) ↑ 越高越好'}, grid:{color:'#f3f4f6'} }
      }
    },
    plugins:[{
      id:'quadrantLines',
      afterDraw(chart) {
        const {ctx, chartArea: {left, right, top, bottom}, scales: {x,y}} = chart;
        if (isNaN(midCpi) || isNaN(midRate)) return;
        const xPos = x.getPixelForValue(Math.min(midCpi, 10));
        const yPos = y.getPixelForValue(midRate);
        ctx.save();
        ctx.strokeStyle = 'rgba(156,163,175,0.5)';
        ctx.lineWidth = 1;
        ctx.setLineDash([4,4]);
        // 竖线
        ctx.beginPath(); ctx.moveTo(xPos, top); ctx.lineTo(xPos, bottom); ctx.stroke();
        // 横线
        ctx.beginPath(); ctx.moveTo(left, yPos); ctx.lineTo(right, yPos); ctx.stroke();
        ctx.setLineDash([]);
        // 象限标签
        ctx.font = '11px sans-serif';
        ctx.fillStyle = '#10b981'; ctx.textAlign = 'left';
        ctx.fillText('★ 爆款', left + 6, top + 14);
        ctx.fillStyle = '#ef4444'; ctx.textAlign = 'right';
        ctx.fillText('✕ 低效', right - 6, bottom - 6);
        ctx.fillStyle = '#f59e0b'; ctx.textAlign = 'right';
        ctx.fillText('高成本高回报', right - 6, top + 14);
        ctx.fillStyle = '#8b5cf6'; ctx.textAlign = 'left';
        ctx.fillText('低成本低回报', left + 6, bottom - 6);
        ctx.restore();
      }
    }]
  });
}

// ============ 款式系列 ============
const SERIES_COLORS = ['#6366f1','#10b981','#f59e0b','#ef4444','#8b5cf6','#ec4899'];
let seriesCostChart, seriesEffChart, seriesTrendChart;
function classifySeries(styleCode) {
  const utri = ['495','0002A1','0001A1','0003A2','493AX','492'];
  // 合并款号别名
  const norm = styleCode === '001A1' || styleCode === '000A1' ? '0001A1' : styleCode;
  return utri.includes(norm) ? 'UTRI系列' : '中台系列';
}
function renderSeriesCharts(rows) {
  const map = {};
  rows.forEach(r => {
    const s = classifySeries(r['款式编号']);
    if (!map[s]) map[s] = { cost: 0, inter: 0, clicks: 0, posts: 0, styles: new Set() };
    map[s].cost += r['总费用']||0;
    map[s].inter += r['总互动']||0;
    map[s].clicks += r['点击量']||0;
    map[s].posts += 1;
    map[s].styles.add(r['款式编号']);
  });
  const arr = Object.entries(map).map(([name,d]) => ({name,...d, styleCount: d.styles.size}))
    .sort((a,b) => b.cost - a.cost);
  const names = arr.map(x=>x.name);
  const costs = arr.map(x=>Math.round(x.cost*10)/10);
  const rates = arr.map(x=>x.clicks>0?Math.round(x.inter/x.clicks*10000)/100:0);
  const cpis = arr.map(x=>x.inter>0?Math.round(x.cost/x.inter*10000)/10000:0);
  const posts = arr.map(x=>x.posts);

  if (seriesCostChart) seriesCostChart.destroy();
  seriesCostChart = new Chart(document.getElementById('seriesCost'), {
    type:'doughnut',
    data:{ labels:names, datasets:[{ data:costs, backgroundColor:SERIES_COLORS, borderWidth:2, borderColor:'#fff', borderRadius:4 }]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ title:{display:true,text:'各系列费用占比',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'right',labels:{padding:10}} }
    }
  });

  if (seriesEffChart) seriesEffChart.destroy();
  seriesEffChart = new Chart(document.getElementById('seriesEff'), {
    type:'bar',
    data:{ labels:names, datasets:[
      { label:'互动率(%)', data:rates, backgroundColor:'rgba(16,185,129,0.85)', borderRadius:4, barThickness:28, yAxisID:'y' },
      { label:'单互动成本(元)', data:cpis, backgroundColor:'rgba(239,68,68,0.85)', borderRadius:4, barThickness:28, yAxisID:'y1' }
    ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ title:{display:true,text:'各系列效率对比',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'bottom'} },
      scales:{ x:{grid:{display:false}}, y:{position:'left',grid:{color:'#f3f4f6'},ticks:{callback:v=>v+'%'}}, y1:{position:'right',grid:{display:false},ticks:{callback:v=>'¥'+v}} }
    }
  });

  // 月度趋势堆叠
  if (seriesTrendChart) seriesTrendChart.destroy();
  const labels = ALL_DATA.monthly.labels.map(l => l.replace('-', '年') + '月');
  const datasets = names.map((name, i) => {
    const monthlyData = ALL_DATA.monthly.labels.map(label => {
      const monthRows = rows.filter(r => r['投放日期'].startsWith(label.slice(0,7)));
      let sum = 0;
      monthRows.forEach(r => {
        if (classifySeries(r['款式编号']) === name) sum += r['总费用']||0;
      });
      return Math.round(sum*10)/10;
    });
    return { label:name, data:monthlyData, backgroundColor:SERIES_COLORS[i%SERIES_COLORS.length], borderRadius:0, barThickness:24 };
  });
  seriesTrendChart = new Chart(document.getElementById('seriesTrend'), {
    type:'bar',
    data:{ labels, datasets },
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ title:{display:true,text:'月度费用结构（各系列）',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'bottom'} },
      scales:{ x:{stacked:true,grid:{display:false}}, y:{stacked:true,grid:{color:'#f3f4f6'},ticks:{callback:v=>'¥'+v.toLocaleString()}} }
    }
  });
}

// ============ 款式 ============
let styleCostChart, styleScatterChart;
function renderStyleCharts(rows) {
  const map = {};
  rows.forEach(r => {
    const code = r['款式编号'];
    if (!map[code]) map[code] = { cost:0, inter:0, clicks:0, posts:0 };
    map[code].cost += r['总费用']||0;
    map[code].inter += r['总互动']||0;
    map[code].clicks += r['点击量']||0;
    map[code].posts += 1;
  });
  const arr = Object.entries(map).map(([name,d]) => ({name,...d}))
    .sort((a,b) => b.cost - a.cost);
  const names = arr.map(x=>x.name);
  const costs = arr.map(x=>Math.round(x.cost*10)/10);
  const rates = arr.map(x=>x.clicks>0?x.inter/x.clicks*100:0);
  const cpis = arr.map(x=>x.inter>0?x.cost/x.inter:999);
  const maxCost = Math.max(...costs, 1);

  if (styleCostChart) styleCostChart.destroy();
  styleCostChart = new Chart(document.getElementById('styleCost'), {
    type:'bar',
    data:{ labels:names, datasets:[{ label:'费用(元)', data:costs, backgroundColor:'rgba(99,102,241,0.85)', borderRadius:3, barThickness:20 }]},
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, title:{display:true,text:'款式投放费用排行',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}} },
      scales:{ x:{grid:{color:'#f3f4f6'},ticks:{callback:v=>'¥'+v.toLocaleString()}}, y:{grid:{display:false}} }
    }
  });

  if (styleScatterChart) styleScatterChart.destroy();
  // 计算中位数作为四象限分割线
  const styleSortedCpis = [...cpis].sort((a,b) => a-b);
  const styleSortedRates = [...rates].sort((a,b) => a-b);
  const styleMidCpi = styleSortedCpis[Math.floor(styleSortedCpis.length/2)];
  const styleMidRate = styleSortedRates[Math.floor(styleSortedRates.length/2)];

  // 四象限颜色
  const styleColors = arr.map((x, i) => {
    const isHighRate = rates[i] >= styleMidRate;
    const isLowCost = cpis[i] <= styleMidCpi;
    if (isHighRate && isLowCost) return { bg: 'rgba(16,185,129,0.6)', border: '#10b981' };
    if (isHighRate && !isLowCost) return { bg: 'rgba(245,158,11,0.6)', border: '#f59e0b' };
    if (!isHighRate && !isLowCost) return { bg: 'rgba(239,68,68,0.6)', border: '#ef4444' };
    return { bg: 'rgba(59,130,246,0.5)', border: '#3b82f6' };
  });

  styleScatterChart = new Chart(document.getElementById('styleScatter'), {
    type:'scatter',
    data:{ datasets:[{
      label:'款式',
      data: arr.map((x,i) => ({
        x: Math.min(cpis[i], 10),
        y: Math.round(rates[i]*100)/100,
        r: 5 + (costs[i]/maxCost)*12,
        name: x.name,
        posts: x.posts,
      })),
      backgroundColor: styleColors.map(c => c.bg),
      borderColor: styleColors.map(c => c.border),
      borderWidth: 1.5,
    }]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false}, title:{display:true,text:'款式效率矩阵（气泡=费用，颜色=象限）',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}},
        tooltip:{ callbacks:{ label:ctx=>{ const d=ctx.raw; return [d.name,'笔记: '+d.posts+'条','互动率: '+d.y+'%','单互动成本: ¥'+Math.round(d.x*10000)/10000]; } }}
      },
      scales:{
        x:{ title:{display:true,text:'单互动成本(元) ← 越左越优'}, grid:{color:'#f3f4f6'} },
        y:{ title:{display:true,text:'互动率(%) ↑ 越高越好'}, grid:{color:'#f3f4f6'} }
      }
    },
    plugins:[{
      id:'quadrantLinesStyle',
      afterDraw(chart) {
        const {ctx, chartArea: {left, right, top, bottom}, scales: {x,y}} = chart;
        if (isNaN(styleMidCpi) || isNaN(styleMidRate)) return;
        const xPos = x.getPixelForValue(Math.min(styleMidCpi, 10));
        const yPos = y.getPixelForValue(styleMidRate);
        ctx.save();
        ctx.strokeStyle = 'rgba(156,163,175,0.5)';
        ctx.lineWidth = 1;
        ctx.setLineDash([4,4]);
        ctx.beginPath(); ctx.moveTo(xPos, top); ctx.lineTo(xPos, bottom); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(left, yPos); ctx.lineTo(right, yPos); ctx.stroke();
        ctx.setLineDash([]);
        ctx.font = '11px sans-serif';
        ctx.fillStyle = '#10b981'; ctx.textAlign = 'left';
        ctx.fillText('★ 爆款', left + 6, top + 14);
        ctx.fillStyle = '#ef4444'; ctx.textAlign = 'right';
        ctx.fillText('✕ 低效', right - 6, bottom - 6);
        ctx.fillStyle = '#f59e0b'; ctx.textAlign = 'right';
        ctx.fillText('高成本高回报', right - 6, top + 14);
        ctx.fillStyle = '#3b82f6'; ctx.textAlign = 'left';
        ctx.fillText('低成本低回报', left + 6, bottom - 6);
        ctx.restore();
      }
    }]
  });
}

// ============ 平台 ============
let platCostChart, platEffChart;
function renderPlatformCharts(rows) {
  const map = {};
  rows.forEach(r => {
    const p = r['平台'];
    if (!map[p]) map[p] = { cost:0, inter:0, clicks:0, posts:0 };
    map[p].cost += r['总费用']||0;
    map[p].inter += r['总互动']||0;
    map[p].clicks += r['点击量']||0;
    map[p].posts += 1;
  });
  const arr = Object.entries(map).map(([name,d]) => ({name,...d}))
    .sort((a,b) => b.cost - a.cost);
  const names = arr.map(x=>x.name);
  const costs = arr.map(x=>Math.round(x.cost*10)/10);
  const rates = arr.map(x=>x.clicks>0?x.inter/x.clicks*100:0);
  const cpis = arr.map(x=>x.inter>0?x.cost/x.inter:0);

  if (platCostChart) platCostChart.destroy();
  platCostChart = new Chart(document.getElementById('platCost'), {
    type:'doughnut',
    data:{ labels:names, datasets:[{ data:costs, backgroundColor:PLAT_COLORS, borderWidth:2, borderColor:'#fff', borderRadius:4 }]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ title:{display:true,text:'各平台费用占比',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'right',labels:{padding:10}} }
    }
  });

  if (platEffChart) platEffChart.destroy();
  platEffChart = new Chart(document.getElementById('platEff'), {
    type:'bar',
    data:{ labels:names, datasets:[
      { label:'互动率(%)', data:rates.map(v=>Math.round(v*100)/100), backgroundColor:'rgba(16,185,129,0.85)', borderRadius:4, barThickness:24, yAxisID:'y' },
      { label:'单互动成本(元)', data:cpis.map(v=>Math.round(v*10000)/10000), backgroundColor:'rgba(239,68,68,0.85)', borderRadius:4, barThickness:24, yAxisID:'y1' }
    ]},
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ title:{display:true,text:'各平台效率对比',font:{size:14,weight:'600'},color:'#1a1d21',padding:{bottom:12}}, legend:{position:'bottom'} },
      scales:{ x:{grid:{display:false}}, y:{position:'left',grid:{color:'#f3f4f6'},ticks:{callback:v=>v+'%'}}, y1:{position:'right',grid:{display:false},ticks:{callback:v=>'¥'+v}} }
    }
  });
}

// ============ 爆款笔记榜 ============
// 入围门槛（保守版）
function meetsThreshold(plat, clicks, inter) {
  if (clicks <= 0 || inter <= 0) return false;
  if (plat === '抖音') return clicks >= 1000 && inter >= 50;
  if (plat === '小红书') return clicks >= 300 && inter >= 20;
  return clicks >= 500 && inter >= 30;
}
// 分平台百分位综合分计算
function computeScores(rows) {
  if (rows.length === 0) return [];
  const byPlat = {};
  rows.forEach(r => {
    const p = r['平台'];
    if (!byPlat[p]) byPlat[p] = [];
    byPlat[p].push(r);
  });
  Object.keys(byPlat).forEach(plat => {
    const list = byPlat[plat];
    if (list.length < 2) {
      list.forEach(r => { r.综合分 = 50; });
      return;
    }
    const n = list.length;
    const sortedClick = [...list].sort((a,b) => a['点击量'] - b['点击量']);
    sortedClick.forEach((r, i) => { r._clickPct = (i + 0.5) / n * 100; });
    const sortedRate = [...list].sort((a,b) => (a['总互动']/a['点击量']) - (b['总互动']/b['点击量']));
    sortedRate.forEach((r, i) => { r._ratePct = (i + 0.5) / n * 100; });
    const sortedInter = [...list].sort((a,b) => a['总互动'] - b['总互动']);
    sortedInter.forEach((r, i) => { r._interPct = (i + 0.5) / n * 100; });
    const sortedCpi = [...list].sort((a,b) => (b['总费用']/b['总互动']) - (a['总费用']/a['总互动']));
    sortedCpi.forEach((r, i) => { r._cpiPct = (i + 0.5) / n * 100; });
    list.forEach(r => {
      r.综合分 = r._clickPct * 0.30 + r._ratePct * 0.30 + r._interPct * 0.25 + r._cpiPct * 0.15;
    });
  });
  return rows;
}

function renderTopNotes() {
  if (!filteredData || filteredData.length === 0) return;
  const monthVal = document.getElementById('topMonthFilter').value;
  const sortVal = document.getElementById('topSortFilter').value;

  let rows = filteredData.slice();
  if (monthVal !== 'all') {
    rows = rows.filter(r => r['投放日期'].startsWith(monthVal));
  }
  // 入围门槛
  rows = rows.filter(r => meetsThreshold(r['平台'], r['点击量'], r['总互动']));
  // 计算综合分
  rows = computeScores(rows);

  // 排序
  if (sortVal === 'rate') {
    rows.sort((a, b) => (b['总互动']/b['点击量']) - (a['总互动']/a['点击量']));
  } else if (sortVal === 'cpi') {
    rows.sort((a, b) => {
      const ca = a['总互动'] > 0 ? a['总费用']/a['总互动'] : 999;
      const cb = b['总互动'] > 0 ? b['总费用']/b['总互动'] : 999;
      return ca - cb;
    });
  } else if (sortVal === 'inter') {
    rows.sort((a, b) => b['总互动'] - a['总互动']);
  } else {
    rows.sort((a, b) => b.综合分 - a.综合分);
  }

  const top5 = rows.slice(0, 10);

  // 表头
  const headRow = document.getElementById('topTableHead');
  headRow.innerHTML = '';
  const headers = ['排名', '综合分', '日期', '平台', '系列', '款式', '内容类型', '点击量', '互动率', '互动量', '单互动成本', '笔记'];
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    headRow.appendChild(th);
  });

  // 表体
  const body = document.getElementById('topTableBody');
  body.innerHTML = '';
  top5.forEach((r, i) => {
    const tr = document.createElement('tr');
    const rate = r['点击量'] > 0 ? (r['总互动'] / r['点击量'] * 100).toFixed(1) + '%' : '-';
    const cpi = r['总互动'] > 0 ? '¥' + (r['总费用'] / r['总互动']).toFixed(4) : '-';
    const rankClass = i === 0 ? 'top-rank-1' : i === 1 ? 'top-rank-2' : i === 2 ? 'top-rank-3' : 'top-rank-n';
    const scoreColor = r.综合分 >= 80 ? 'color:#10b981;font-weight:700' : r.综合分 >= 70 ? 'color:#6366f1;font-weight:600' : 'color:#6b7280';

    const cells = [
      `<span class="top-rank ${rankClass}">${i+1}</span>`,
      `<span style="${scoreColor}">${r.综合分.toFixed(1)}</span>`,
      r['投放日期'] || '-',
      r['平台'] || '-',
      r['款式系列'] || '-',
      r['款式编号'] || '-',
      r['内容类型'] || '-',
      Math.round(r['点击量'] || 0).toLocaleString(),
      rate,
      Math.round(r['总互动'] || 0).toLocaleString(),
      cpi,
    ];

    cells.forEach(text => {
      const td = document.createElement('td');
      td.innerHTML = text;
      tr.appendChild(td);
    });

    // 笔记链接
    const td = document.createElement('td');
    if (r['内容链接']) {
      const a = document.createElement('a');
      a.href = r['内容链接'];
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.className = 'top-link';
      a.textContent = '🔗 查看';
      td.appendChild(a);
    } else {
      td.textContent = '-';
    }
    tr.appendChild(td);

    body.appendChild(tr);
  });

  if (top5.length === 0) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 12;
    td.style.textAlign = 'center';
    td.style.color = '#9ca3af';
    td.style.padding = '20px';
    td.textContent = '暂无达标笔记';
    tr.appendChild(td);
    body.appendChild(tr);
  }
}

// ============ 明细表 ============
function renderTable() {
  const search = document.getElementById('tableSearch').value.toLowerCase();
  let rows = filteredData.slice();

  if (search) {
    rows = rows.filter(r =>
      (r['内容类型']||'').toLowerCase().includes(search) ||
      (r['款式编号']||'').toLowerCase().includes(search) ||
      (r['平台']||'').toLowerCase().includes(search) ||
      (r['投放日期']||'').includes(search) ||
      (r['达人姓名']||'').toLowerCase().includes(search) ||
      (r['款式系列']||'').toLowerCase().includes(search)
    );
  }

  // 排序
  if (sortCol >= 0) {
    const col = ALL_DATA.detailColumns[sortCol];
    rows.sort((a,b) => {
      let va = a[col], vb = b[col];
      if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * sortDir;
      return String(va).localeCompare(String(vb), 'zh') * sortDir;
    });
  }

  document.getElementById('tableCount').textContent = rows.length + ' 条记录';

  // 表头
  const headRow = document.getElementById('tableHead');
  headRow.innerHTML = '';
  ALL_DATA.detailColumns.forEach((c, i) => {
    const th = document.createElement('th');
    th.textContent = c + (sortCol === i ? (sortDir > 0 ? ' ↑' : ' ↓') : '');
    th.onclick = () => { sortCol = i; sortDir = -sortDir; renderTable(); };
    headRow.appendChild(th);
  });

  // 表体（取前300条避免卡顿）
  const body = document.getElementById('tableBody');
  body.innerHTML = '';
  const showRows = rows.slice(0, 300);
  showRows.forEach(r => {
    const tr = document.createElement('tr');
    ALL_DATA.detailColumns.forEach(c => {
      const td = document.createElement('td');
      let v = r[c];
      if (c === '互动率') v = v ? v.toFixed(2) + '%' : '-';
      else if (['达人合作费','平台投放费','样品/运费成本','总费用','单互动成本'].includes(c)) v = v ? '¥' + Number(v).toFixed(2) : '-';
      else if (c === '内容链接' && v) {
        const a = document.createElement('a');
        a.href = v;
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        a.textContent = '🔗 查看笔记';
        a.style.cssText = 'display:inline-block;padding:3px 10px;background:#eef2ff;color:#4f46e5;border-radius:4px;font-size:11px;text-decoration:none;white-space:nowrap';
        a.onmouseenter = () => { a.style.background = '#e0e7ff'; };
        a.onmouseleave = () => { a.style.background = '#eef2ff'; };
        td.appendChild(a);
        tr.appendChild(td);
        return;
      }
      else if (typeof v === 'number') v = Math.round(v).toLocaleString();
      else if (v === undefined || v === null || v === '') v = '-';
      td.textContent = v;
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });

  if (rows.length > 300) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = ALL_DATA.detailColumns.length;
    td.style.textAlign = 'center';
    td.style.color = '#9ca3af';
    td.style.padding = '12px';
    td.textContent = `仅显示前300条，共${rows.length}条。使用搜索框缩小范围~`;
    tr.appendChild(td);
    body.appendChild(tr);
  }
}

// ============ 初始化 ============
function init() {
  // 填充平台/系列/款式/月份下拉
  const platSel = document.getElementById('platFilter');
  ALL_DATA.platform.names.forEach(p => {
    const opt = document.createElement('option');
    opt.value = p; opt.textContent = p;
    platSel.appendChild(opt);
  });
  const monthSel = document.getElementById('monthFilter');
  ALL_DATA.monthly.labels.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = m.replace('-', '年') + '月';
    monthSel.appendChild(opt);
  });
  // 爆款榜月份下拉
  const topMonthSel = document.getElementById('topMonthFilter');
  ALL_DATA.monthly.labels.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = m.replace('-', '年') + '月';
    topMonthSel.appendChild(opt);
  });
  const seriesSel = document.getElementById('seriesFilter');
  ALL_DATA.series.names.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    seriesSel.appendChild(opt);
  });
  const styleSel = document.getElementById('styleFilter');
  ALL_DATA.style.names.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    styleSel.appendChild(opt);
  });

  applyFilter();
}

init();
</script>
</body>
</html>
"""

# ============================================================
# 主程序
# ============================================================
if __name__ == '__main__':
    input_file = sys.argv[1] if len(sys.argv) > 1 else r'C:\Users\Administrator\达人投放数据_1-7月.xlsx'
    output_file = sys.argv[2] if len(sys.argv) > 2 else r'C:\Users\Administrator\达人投放看板.html'

    print(f'📊 读取数据：{input_file}')
    df = load_data(input_file)
    print(f'   共 {len(df)} 条记录')

    print(f'🔍 计算分析数据...')
    data = build_chart_data(df)
    print(f'   内容类型：{data["ctype"]["totalTypes"]} 种')
    print(f'   款式数量：{len(data["style"]["names"])} 个')
    print(f'   平台数量：{len(data["platform"]["names"])} 个')

    print(f'🎨 生成看板...')
    generate_html(data, output_file)
    print(f'🎉 完成！输出：{output_file}')
