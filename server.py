"""
实验室样本出入库管理系统 - 后端服务
Flask + SQLite + Excel导出
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

app = Flask(__name__)
CORS(app)

# 数据库路径
DB_DIR = os.path.join(os.path.dirname(__file__), 'data')
DB_PATH = os.path.join(DB_DIR, 'lab.db')
os.makedirs(DB_DIR, exist_ok=True)


# ===== 初始化数据库 =====
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    
    # 样本表
    c.execute('''
        CREATE TABLE IF NOT EXISTS samples (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            样品分子号 TEXT NOT NULL UNIQUE,
            患者姓名 TEXT NOT NULL,
            患者性别 TEXT,
            患者年龄 INTEGER,
            诊断 TEXT,
            样本种类 TEXT,
            冰箱编号 TEXT,
            冰箱层 TEXT,
            收纳盒编号 TEXT,
            格子序号 TEXT,
            入库管数 INTEGER DEFAULT 1,
            当前余量 INTEGER DEFAULT 1,
            操作日期 TEXT,
            补充说明 TEXT,
            创建时间 TEXT DEFAULT (datetime('now','localtime')),
            更新时间 TEXT DEFAULT (datetime('now','localtime'))
        )
    ''')
    
    # 操作记录表（入库/取用/归还日志）
    c.execute('''
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            操作类型 TEXT NOT NULL,
            样品分子号 TEXT,
            患者姓名 TEXT,
            操作数量 INTEGER,
            操作日期 TEXT,
            用途 TEXT,
            操作人 TEXT,
            备注 TEXT,
            创建时间 TEXT DEFAULT (datetime('now','localtime'))
        )
    ''')
    
    conn.commit()
    conn.close()


init_db()


# ===== 获取数据库连接 =====
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def dict_from_row(row):
    """将sqlite3.Row转为dict"""
    if row is None:
        return None
    return dict(row)


# ===== API 路由 =====

@app.route('/')
def index():
    return jsonify({"status": "ok", "message": "实验室样本出入库管理系统 API"})


@app.route('/api/samples', methods=['GET'])
def get_samples():
    """获取所有样本列表"""
    conn = get_db()
    rows = conn.execute('SELECT * FROM samples ORDER BY 更新时间 DESC').fetchall()
    conn.close()
    samples = [dict_from_row(r) for r in rows]
    
    # 增加库位状态字段
    for s in samples:
        if s['当前余量'] <= 0:
            s['库位状态'] = '已取完'
        elif s['当前余量'] < s['入库管数']:
            s['库位状态'] = '部分取走'
        else:
            s['库位状态'] = '有样本'
    
    return jsonify({"samples": samples})


@app.route('/api/samples/<sample_id>', methods=['GET'])
def get_sample(sample_id):
    """获取单个样本详情"""
    conn = get_db()
    row = conn.execute('SELECT * FROM samples WHERE 样品分子号 = ?', (sample_id,)).fetchone()
    conn.close()
    
    if row is None:
        return jsonify({"error": "未找到该样本"}), 404
    
    return jsonify(dict_from_row(row))


@app.route('/api/inbound', methods=['POST'])
def add_inbound():
    """入库登记"""
    data = request.get_json()
    if not data:
        return jsonify({"error": "请提供数据"}), 400
    
    required = ['样品分子号', '患者姓名']
    for field in required:
        if not data.get(field):
            return jsonify({"error": f"缺少必填字段: {field}"}), 400
    
    conn = get_db()
    
    # 检查是否已存在
    existing = conn.execute(
        'SELECT * FROM samples WHERE 样品分子号 = ?', 
        (data['样品分子号'],)
    ).fetchone()
    
    if existing:
        # 已存在：增加库存
        old_count = existing['入库管数']
        old_remaining = existing['当前余量']
        new_count = old_count + int(data.get('入库管数', 1))
        new_remaining = old_remaining + int(data.get('入库管数', 1))
        
        conn.execute('''
            UPDATE samples SET 
                入库管数 = ?, 当前余量 = ?, 更新时间 = datetime('now','localtime')
            WHERE 样品分子号 = ?
        ''', (new_count, new_remaining, data['样品分子号']))
    else:
        # 新增
        count = int(data.get('入库管数', 1))
        conn.execute('''
            INSERT INTO samples 
            (样品分子号, 患者姓名, 患者性别, 患者年龄, 诊断, 样本种类,
             冰箱编号, 冰箱层, 收纳盒编号, 格子序号, 入库管数, 当前余量, 操作日期, 补充说明)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data['样品分子号'], data['患者姓名'], data.get('患者性别', ''),
            data.get('患者年龄'), data.get('诊断', ''), data.get('样本种类', ''),
            data.get('冰箱编号', ''), data.get('冰箱层', ''),
            data.get('收纳盒编号', ''), data.get('格子序号', ''),
            count, count, data.get('操作日期', datetime.now().strftime('%Y-%m-%d')),
            data.get('补充说明', '')
        ))
    
    # 写操作记录
    conn.execute('''
        INSERT INTO records (操作类型, 样品分子号, 患者姓名, 操作数量, 操作日期, 备注)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', ('入库', data['样品分子号'], data['患者姓名'], 
           int(data.get('入库管数', 1)),
           data.get('操作日期', datetime.now().strftime('%Y-%m-%d')),
           data.get('补充说明', '')))
    
    conn.commit()
    conn.close()
    
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
    
    conn = get_db()
    
    # 检查库存
    sample = conn.execute(
        'SELECT * FROM samples WHERE 样品分子号 = ?', 
        (sample_id,)
    ).fetchone()
    
    if sample is None:
        conn.close()
        return jsonify({"error": "未找到该样本"}), 404
    
    remaining = sample['当前余量']
    if remaining < count:
        conn.close()
        return jsonify({"error": f"库存不足！当前余量 {remaining} 管，需取 {count} 管"}), 400
    
    # 更新余量
    conn.execute('''
        UPDATE samples SET 
            当前余量 = 当前余量 - ?,
            更新时间 = datetime('now','localtime')
        WHERE 样品分子号 = ?
    ''', (count, sample_id))
    
    # 写操作记录
    conn.execute('''
        INSERT INTO records (操作类型, 样品分子号, 患者姓名, 操作数量, 操作日期, 用途, 操作人)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', ('取用', sample_id, sample['患者姓名'], count,
           data.get('操作日期', datetime.now().strftime('%Y-%m-%d')),
           data.get('用途', ''), data.get('操作人', '')))
    
    conn.commit()
    conn.close()
    
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
    
    conn = get_db()
    
    # 检查样本是否存在
    sample = conn.execute(
        'SELECT * FROM samples WHERE 样品分子号 = ?', 
        (sample_id,)
    ).fetchone()
    
    if sample is None:
        conn.close()
        return jsonify({"error": "未找到该样本"}), 404
    
    # 更新余量
    conn.execute('''
        UPDATE samples SET 
            当前余量 = 当前余量 + ?,
            更新时间 = datetime('now','localtime')
        WHERE 样品分子号 = ?
    ''', (count, sample_id))
    
    # 写操作记录
    conn.execute('''
        INSERT INTO records (操作类型, 样品分子号, 患者姓名, 操作数量, 操作日期, 操作人)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', ('归还', sample_id, sample['患者姓名'], count,
           data.get('操作日期', datetime.now().strftime('%Y-%m-%d')),
           data.get('操作人', '')))
    
    conn.commit()
    conn.close()
    
    return jsonify({"status": "success", "message": "归还成功"})


@app.route('/api/records', methods=['GET'])
def get_records():
    """获取操作记录"""
    conn = get_db()
    rows = conn.execute('SELECT * FROM records ORDER BY 创建时间 DESC LIMIT 500').fetchall()
    conn.close()
    return jsonify({"records": [dict_from_row(r) for r in rows]})


@app.route('/')
def serve_index():
    return send_from_directory('templates', 'index.html')


@app.route('/api/export/excel')
def export_excel():
    """导出Excel报表"""
    conn = get_db()
    
    # ---- 样本库存表 ----
    samples = conn.execute('SELECT * FROM samples ORDER BY 更新时间 DESC').fetchall()
    
    # ---- 操作记录表 ----
    records = conn.execute('SELECT * FROM records ORDER BY 创建时间 DESC').fetchall()
    
    conn.close()
    
    wb = Workbook()
    
    # ===== Sheet 1: 样本库存 =====
    ws1 = wb.active
    ws1.title = "样本库存"
    
    # 样式
    header_font = Font(bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="1B3A5C", end_color="1B3A5C", fill_type="solid")
    header_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
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
        remaining = s['当前余量']
        if remaining <= 0:
            status = '已取完'
        elif remaining < s['入库管数']:
            status = '部分取走'
        else:
            status = '有样本'
        
        values = [
            s['样品分子号'], s['患者姓名'], s['患者性别'], s['患者年龄'],
            s['诊断'], s['样本种类'], s['冰箱编号'], s['冰箱层'],
            s['收纳盒编号'], s['格子序号'], s['入库管数'], s['当前余量'],
            status, s['操作日期'], s['补充说明']
        ]
        
        for col, val in enumerate(values, 1):
            cell = ws1.cell(row=row_idx, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')
    
    # 自动列宽
    for col in range(1, len(headers) + 1):
        max_len = len(str(ws1.cell(row=1, column=col).value))
        for row in range(2, len(samples) + 2):
            cell_val = str(ws1.cell(row=row, column=col).value or '')
            max_len = max(max_len, len(cell_val))
        ws1.column_dimensions[chr(64 + col)].width = min(max_len + 4, 30)
    
    # ===== Sheet 2: 操作记录 =====
    ws2 = wb.create_sheet("操作记录")
    
    record_headers = ['ID', '操作类型', '样品分子号', '患者姓名', '操作数量',
                      '操作日期', '用途', '操作人', '备注', '创建时间']
    
    for col, h in enumerate(record_headers, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    for row_idx, r in enumerate(records, 2):
        values = [
            r['id'], r['操作类型'], r['样品分子号'], r['患者姓名'],
            r['操作数量'], r['操作日期'], r['用途'], r['操作人'],
            r['备注'], r['创建时间']
        ]
        for col, val in enumerate(values, 1):
            cell = ws2.cell(row=row_idx, column=col, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center')
    
    for col in range(1, len(record_headers) + 1):
        max_len = len(str(ws2.cell(row=1, column=col).value))
        for row in range(2, len(records) + 2):
            cell_val = str(ws2.cell(row=row, column=col).value or '')
            max_len = max(max_len, len(cell_val))
        ws2.column_dimensions[chr(64 + col)].width = min(max_len + 4, 30)
    
    # 输出到内存
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"实验室样本库存_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5051))
    print(f"🚀 实验室管理系统启动中... http://0.0.0.0:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
