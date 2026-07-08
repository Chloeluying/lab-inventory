"""
实验室样本出入库管理系统 - Vercel Serverless版本
使用内存SQLite + 飞书bitable作为持久化存储
"""

import os
import sqlite3
import json
from datetime import datetime
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from io import BytesIO
import requests

app = Flask(__name__)
CORS(app)

# ===== 飞书API配置 =====
FEISHU_APP_ID = os.environ.get('FEISHU_APP_ID', '')
FEISHU_APP_SECRET = os.environ.get('FEISHU_APP_SECRET', '')
FEISHU_BITABLE_APP_TOKEN = 'H81bbGvwhaPW9WshW4wcuAu4nob'
FEISHU_TABLE_ID = 'tblJmPPRcnqtRoeo'


def get_feishu_token():
    """获取飞书tenant_access_token"""
    r = requests.post('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', json={
        'app_id': FEISHU_APP_ID,
        'app_secret': FEISHU_APP_SECRET
    })
    data = r.json()
    if data.get('code') == 0:
        return data['tenant_access_token']
    return None


def feishu_get_records():
    """从飞书bitable获取所有记录"""
    token = get_feishu_token()
    if not token:
        return []
    
    headers = {'Authorization': f'Bearer {token}'}
    records = []
    page_token = None
    
    while True:
        url = f'https://open.feishu.cn/open-apis/bitable/v1/apps/{FEISHU_BITABLE_APP_TOKEN}/tables/{FEISHU_TABLE_ID}/records?page_size=500'
        if page_token:
            url += f'&page_token={page_token}'
        
        r = requests.get(url, headers=headers)
        data = r.json()
        
        if data.get('code') != 0:
            break
        
        items = data.get('data', {}).get('items', [])
        records.extend(items)
        
        if not data.get('data', {}).get('has_more'):
            break
        page_token = data.get('data', {}).get('page_token')
    
    return records


def feishu_add_record(fields):
    """向飞书bitable添加记录"""
    token = get_feishu_token()
    if not token:
        return False
    
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    
    r = requests.post(
        f'https://open.feishu.cn/open-apis/bitable/v1/apps/{FEISHU_BITABLE_APP_TOKEN}/tables/{FEISHU_TABLE_ID}/records',
        headers=headers,
        json={'fields': fields}
    )
    
    return r.json().get('code') == 0


def feishu_update_record(record_id, fields):
    """更新飞书bitable记录"""
    token = get_feishu_token()
    if not token:
        return False
    
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    
    r = requests.put(
        f'https://open.feishu.cn/open-apis/bitable/v1/apps/{FEISHU_BITABLE_APP_TOKEN}/tables/{FEISHU_TABLE_ID}/records/{record_id}',
        headers=headers,
        json={'fields': fields}
    )
    
    return r.json().get('code') == 0


def records_to_samples(records):
    """将飞书记录转为标准样本格式"""
    samples = []
    for rec in records:
        f = rec.get('fields', {})
        record_id = rec.get('record_id', '')
        
        sample_type = f.get('样本类型', f.get('样本种类', ''))
        if isinstance(sample_type, list):
            sample_type = sample_type[0] if sample_type else ''
        
        diagnosis = f.get('诊断', '')
        if isinstance(diagnosis, list):
            diagnosis = diagnosis[0] if diagnosis else ''
        
        fridge = f.get('冰箱编号', '')
        if isinstance(fridge, list):
            fridge = fridge[0] if fridge else ''
        
        layer = f.get('冰箱层（从上往下）', '')
        if isinstance(layer, list):
            layer = layer[0] if layer else ''
        
        gender = f.get('患者性别', '')
        if isinstance(gender, list):
            gender = gender[0] if gender else ''
        
        op_type = f.get('操作类型', '')
        if isinstance(op_type, list):
            op_type = op_type[0] if op_type else ''
        
        status = f.get('库位状态', '')
        if isinstance(status, list):
            status = status[0] if status else ''
        
        total = int(f.get('入库管数', 0) or 0)
        remaining = int(f.get('当前余量', 0) or 0)
        
        samples.append({
            'record_id': record_id,
            '患者姓名': f.get('患者姓名', ''),
            '患者性别': gender,
            '患者年龄': int(f.get('患者年龄', 0) or 0),
            '诊断': diagnosis,
            '样品分子号': f.get('样品分子号', ''),
            '样本种类': sample_type,
            '冰箱编号': fridge,
            '冰箱层': str(layer),
            '收纳盒编号': f.get('收纳盒编号', ''),
            '格子序号': f.get('格子序号', ''),
            '入库管数': total,
            '当前余量': remaining,
            '库位状态': status if status else ('有样本' if remaining > 0 else '已取完'),
            '操作日期': f.get('日期', ''),
            '补充说明': f.get('补充说明', ''),
            '操作类型': op_type
        })
    
    return samples


# ===== API 路由 =====

@app.route('/')
def serve_index():
    return send_from_directory('templates', 'index.html')


@app.route('/api/samples', methods=['GET'])
def get_samples():
    """获取所有样本（从飞书bitable）"""
    records = feishu_get_records()
    samples = records_to_samples(records)
    return jsonify({"samples": samples})


@app.route('/api/inbound', methods=['POST'])
def add_inbound():
    """入库登记"""
    data = request.get_json()
    if not data:
        return jsonify({"error": "请提供数据"}), 400
    
    sample_id = data.get('样品分子号')
    name = data.get('患者姓名')
    if not sample_id or not name:
        return jsonify({"error": "缺少必填字段"}), 400
    
    count = int(data.get('入库管数', 1))
    
    # 检查是否已存在
    records = feishu_get_records()
    existing = [r for r in records if r.get('fields', {}).get('样品分子号') == sample_id]
    
    if existing:
        # 已存在：更新库存
        rec = existing[0]
        rid = rec['record_id']
        old_total = int(rec['fields'].get('入库管数', 0) or 0)
        old_remaining = int(rec['fields'].get('当前余量', 0) or 0)
        
        feishu_update_record(rid, {
            '入库管数': old_total + count,
            '当前余量': old_remaining + count,
            '库位状态': '有样本' if old_remaining + count > 0 else '全部取走'
        })
    else:
        # 新增
        feishu_add_record({
            '样品分子号': sample_id,
            '患者姓名': name,
            '患者性别': data.get('患者性别', ''),
            '患者年龄': int(data.get('患者年龄', 0) or 0),
            '诊断': data.get('诊断', ''),
            '样本种类': data.get('样本种类', ''),
            '冰箱编号': data.get('冰箱编号', ''),
            '冰箱层（从上往下）': data.get('冰箱层', ''),
            '收纳盒编号': data.get('收纳盒编号', ''),
            '格子序号': data.get('格子序号', ''),
            '入库管数': count,
            '当前余量': count,
            '库位状态': '有样本',
            '操作类型': '入库',
            '日期': int(datetime.now().timestamp() * 1000),
            '补充说明': data.get('补充说明', '')
        })
    
    return jsonify({"status": "success", "message": "入库成功"})


@app.route('/api/outbound', methods=['POST'])
def add_outbound():
    """取用登记"""
    data = request.get_json()
    if not data:
        return jsonify({"error": "请提供数据"}), 400
    
    sample_id = data.get('样品分子号')
    count = int(data.get('操作数量', 1))
    
    if not sample_id:
        return jsonify({"error": "缺少样品分子号"}), 400
    
    records = feishu_get_records()
    existing = [r for r in records if r.get('fields', {}).get('样品分子号') == sample_id]
    
    if not existing:
        return jsonify({"error": "未找到该样本"}), 404
    
    rec = existing[0]
    rid = rec['record_id']
    remaining = int(rec['fields'].get('当前余量', 0) or 0)
    total = int(rec['fields'].get('入库管数', 0) or 0)
    
    if remaining < count:
        return jsonify({"error": f"库存不足！当前余量 {remaining} 管"}), 400
    
    new_remaining = remaining - count
    if new_remaining <= 0:
        status = '全部取走'
    elif new_remaining < total:
        status = '部分取走'
    else:
        status = '有样本'
    
    feishu_update_record(rid, {
        '当前余量': new_remaining,
        '库位状态': status
    })
    
    return jsonify({"status": "success", "message": "取用成功"})


@app.route('/api/return', methods=['POST'])
def add_return():
    """归还登记"""
    data = request.get_json()
    if not data:
        return jsonify({"error": "请提供数据"}), 400
    
    sample_id = data.get('样品分子号')
    count = int(data.get('操作数量', 1))
    
    if not sample_id:
        return jsonify({"error": "缺少样品分子号"}), 400
    
    records = feishu_get_records()
    existing = [r for r in records if r.get('fields', {}).get('样品分子号') == sample_id]
    
    if not existing:
        return jsonify({"error": "未找到该样本"}), 404
    
    rec = existing[0]
    rid = rec['record_id']
    old_remaining = int(rec['fields'].get('当前余量', 0) or 0)
    total = int(rec['fields'].get('入库管数', 0) or 0)
    
    new_remaining = old_remaining + count
    status = '有样本'
    
    feishu_update_record(rid, {
        '当前余量': new_remaining,
        '库位状态': status
    })
    
    return jsonify({"status": "success", "message": "归还成功"})


@app.route('/api/export/excel')
def export_excel():
    """导出Excel报表"""
    records = feishu_get_records()
    samples = records_to_samples(records)
    
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "样本库存"
    
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="1B3A5C", end_color="1B3A5C", fill_type="solid")
    header_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    headers = ['样品分子号', '患者姓名', '患者性别', '患者年龄', '诊断', '样本种类',
               '冰箱编号', '冰箱层', '收纳盒编号', '格子序号', '入库管数', '当前余量',
               '库位状态', '操作日期', '补充说明']
    
    for col, h in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    for row_idx, s in enumerate(samples, 2):
        values = [
            s['样品分子号'], s['患者姓名'], s['患者性别'], s['患者年龄'],
            s['诊断'], s['样本种类'], s['冰箱编号'], s['冰箱层'],
            s['收纳盒编号'], s['格子序号'], s['入库管数'], s['当前余量'],
            s['库位状态'], s['操作日期'], s['补充说明']
        ]
        for col, val in enumerate(values, 1):
            cell = ws1.cell(row=row_idx, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')
    
    for col in range(1, len(headers) + 1):
        max_len = len(str(ws1.cell(row=1, column=col).value))
        for row in range(2, len(samples) + 2):
            cell_val = str(ws1.cell(row=row, column=col).value or '')
            max_len = max(max_len, len(cell_val))
        ws1.column_dimensions[chr(64 + col)].width = min(max_len + 4, 30)
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"实验室样本库存_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )


# 本地开发用
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5051))
    app.run(host='0.0.0.0', port=port, debug=False)
