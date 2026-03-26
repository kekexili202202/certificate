from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import sqlite3
import base64
from datetime import datetime
import pandas as pd  # 新增：用于读取 Excel

app = Flask(__name__, static_url_path='', static_folder='.')
CORS(app)

# 管理员密钥（可以改成你自己的密码）
ADMIN_KEY = "admin123456"


# 初始化数据库
def init_db():
    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    # 选手数据表
    c.execute('''
        CREATE TABLE IF NOT EXISTS participants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            region TEXT NOT NULL,
            phone TEXT NOT NULL,
            organization TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 模板数据表
    c.execute('''
        CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cert_type TEXT NOT NULL,
            region TEXT NOT NULL,
            filename TEXT NOT NULL,
            pdf_data TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    conn.commit()
    conn.close()
    print("数据库初始化完成")


init_db()


# ==================== 选手数据接口 ====================

@app.route('/api/participants', methods=['GET'])
def get_participants():
    """获取所有选手数据"""
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('SELECT id, name, region, phone, organization FROM participants ORDER BY id')
    rows = c.fetchall()
    conn.close()

    participants = []
    for row in rows:
        participants.append({
            'id': row[0],
            '姓名': row[1],
            '赛区': row[2],
            '手机号': row[3],
            '所在单位': row[4]
        })
    return jsonify(participants)


@app.route('/api/participants/upload', methods=['POST'])
def upload_participants():
    """上传选手数据"""
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    data = request.json
    participants = data.get('participants', [])

    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    # 清空原有数据
    c.execute('DELETE FROM participants')

    # 插入新数据
    for p in participants:
        c.execute('''
            INSERT INTO participants (name, region, phone, organization)
            VALUES (?, ?, ?, ?)
        ''', (p.get('姓名', ''), p.get('赛区', ''), p.get('手机号', ''), p.get('所在单位', '')))

    conn.commit()
    conn.close()

    return jsonify({'success': True, 'count': len(participants)})


@app.route('/api/participants/query', methods=['POST'])
def query_participant():
    """查询选手信息"""
    data = request.json
    name = data.get('name', '').strip()
    region = data.get('region', '').strip()
    phone = data.get('phone', '').strip()
    organization = data.get('organization', '').strip()

    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('''
        SELECT name, region, phone, organization FROM participants
        WHERE name = ? AND region = ? AND phone = ? AND organization = ?
    ''', (name, region, phone, organization))

    row = c.fetchone()
    conn.close()

    if row:
        return jsonify({
            'found': True,
            'participant': {
                '姓名': row[0],
                '赛区': row[1],
                '手机号': row[2],
                '所在单位': row[3]
            }
        })
    else:
        return jsonify({'found': False})


# ==================== 模板管理接口 ====================

@app.route('/api/templates', methods=['GET'])
def get_templates():
    """获取所有模板"""
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('SELECT cert_type, region, filename, pdf_data FROM templates')
    rows = c.fetchall()
    conn.close()

    templates = {
        'participation': {},
        'preliminary': {},
        'final': {}
    }

    for row in rows:
        cert_type = row[0]
        region = row[1]
        templates[cert_type][region] = {
            'filename': row[2],
            'pdf_data': row[3]
        }

    return jsonify(templates)


@app.route('/api/templates/upload', methods=['POST'])
def upload_template():
    """上传模板"""
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    data = request.json
    cert_type = data.get('cert_type')
    region = data.get('region')
    filename = data.get('filename')
    pdf_data = data.get('pdf_data')

    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    # 删除旧的同类型同赛区模板
    c.execute('DELETE FROM templates WHERE cert_type = ? AND region = ?', (cert_type, region))

    # 插入新模板
    c.execute('''
        INSERT INTO templates (cert_type, region, filename, pdf_data)
        VALUES (?, ?, ?, ?)
    ''', (cert_type, region, filename, pdf_data))

    conn.commit()
    conn.close()

    return jsonify({'success': True})


@app.route('/api/templates/delete', methods=['DELETE'])
def delete_template():
    """删除模板"""
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    data = request.json
    cert_type = data.get('cert_type')
    region = data.get('region')

    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('DELETE FROM templates WHERE cert_type = ? AND region = ?', (cert_type, region))
    conn.commit()
    conn.close()

    return jsonify({'success': True})


@app.route('/api/templates/clear', methods=['DELETE'])
def clear_templates():
    """清空所有模板"""
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('DELETE FROM templates')
    conn.commit()
    conn.close()

    return jsonify({'success': True})


# ==================== 新增：从服务器读取 Excel 接口 ====================

@app.route('/api/load-excel', methods=['GET'])
def load_excel():
    """从服务器读取 Excel 文件"""
    try:
        # 读取项目文件夹中的 Excel 文件
        df = pd.read_excel('法律英语大赛网站测试.xlsx')

        # 转换成 JSON 格式
        participants = []
        for _, row in df.iterrows():
            participants.append({
                '姓名': str(row.get('姓名', '')),
                '赛区': str(row.get('赛区', '')),
                '手机号': str(row.get('手机号', '')),
                '所在单位': str(row.get('所在单位', ''))
            })

        return jsonify({'success': True, 'data': participants, 'count': len(participants)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


# ==================== 页面路由 ====================

@app.route('/')
def index():
    """查询页面（分享给用户）"""
    return send_from_directory('.', 'query.html')


@app.route('/admin')
def admin():
    """管理员页面（你自己用）"""
    return send_from_directory('.', 'admin.html')


@app.route('/<path:path>')
def serve_static(path):
    """提供静态文件"""
    return send_from_directory('.', path)


if __name__ == '__main__':
    import os

    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)