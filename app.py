from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import sqlite3
import base64
from datetime import datetime
import os

app = Flask(__name__, static_url_path='', static_folder='.')
CORS(app)

# 管理员密钥
ADMIN_KEY = "admin123456"

# 赛区编号映射
REGION_CODE = {
    "华北及东北赛区": "01",
    "华南赛区": "02",
    "华东赛区": "03",
    "西南赛区": "04",
    "西北赛区": "05"
}

CODE_TO_REGION = {v: k for k, v in REGION_CODE.items()}


def init_db():
    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='participants'")
    table_exists = c.fetchone()

    if not table_exists:
        c.execute('''
            CREATE TABLE participants (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                region TEXT NOT NULL,
                region_code TEXT NOT NULL,
                phone TEXT NOT NULL,
                organization TEXT NOT NULL,
                certificate_number TEXT DEFAULT '',
                award_level TEXT DEFAULT '',
                cert_type TEXT DEFAULT '',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        print("创建 participants 表")
    else:
        c.execute("PRAGMA table_info(participants)")
        columns = [col[1] for col in c.fetchall()]

        if 'region_code' not in columns:
            c.execute("ALTER TABLE participants ADD COLUMN region_code TEXT DEFAULT ''")
        if 'certificate_number' not in columns:
            c.execute("ALTER TABLE participants ADD COLUMN certificate_number TEXT DEFAULT ''")
        if 'award_level' not in columns:
            c.execute("ALTER TABLE participants ADD COLUMN award_level TEXT DEFAULT ''")
        if 'cert_type' not in columns:
            c.execute("ALTER TABLE participants ADD COLUMN cert_type TEXT DEFAULT ''")
        if 'created_at' not in columns:
            c.execute("ALTER TABLE participants ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

    c.execute('''
        CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cert_type TEXT NOT NULL,
            region TEXT NOT NULL,
            award_level TEXT DEFAULT '',
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
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute(
        'SELECT id, name, region, phone, organization, certificate_number, award_level, cert_type FROM participants ORDER BY id')
    rows = c.fetchall()
    conn.close()

    participants = []
    for row in rows:
        participants.append({
            'id': row[0],
            '姓名': row[1],
            '赛区': row[2],
            '手机号': row[3],
            '所在单位': row[4],
            '证书编号': row[5] if row[5] else '',
            '奖项等级': row[6] if row[6] else '',
            '证书类型': row[7] if row[7] else ''
        })
    return jsonify(participants)


@app.route('/api/participants/upload', methods=['POST'])
def upload_participants():
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    data = request.json
    sheets_data = data.get('sheets', {})

    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    c.execute('DELETE FROM participants')

    total_count = 0

    for cert_type, participants in sheets_data.items():
        for p in participants:
            name = p.get('姓名', '')
            region = p.get('赛区', '')
            region_code = REGION_CODE.get(region, '00')
            phone = p.get('手机号', '')
            organization = p.get('所在单位', '')
            certificate_number = p.get('证书编号', '')
            award_level = p.get('奖项等级', '')

            c.execute('''
                INSERT INTO participants (name, region, region_code, phone, organization, certificate_number, award_level, cert_type)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (name, region, region_code, phone, organization, certificate_number, award_level, cert_type))
            total_count += 1

    conn.commit()
    conn.close()

    return jsonify({'success': True, 'count': total_count})


@app.route('/api/participants/clear', methods=['DELETE', 'POST'])
def clear_participants():
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('DELETE FROM participants')
    conn.commit()
    conn.close()

    return jsonify({'success': True, 'count': 0})


@app.route('/api/participants/query', methods=['POST'])
def query_participant():
    data = request.json
    name = data.get('name', '').strip()
    region = data.get('region', '').strip()
    phone = data.get('phone', '').strip()
    organization = data.get('organization', '').strip()

    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('''
        SELECT name, region, phone, organization, certificate_number FROM participants
        WHERE name = ? AND region = ? AND phone = ? AND organization = ? AND (cert_type = '' OR cert_type = 'participation')
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
                '所在单位': row[3],
                '证书编号': row[4] if row[4] else ''
            }
        })
    else:
        return jsonify({'found': False})


@app.route('/api/participants/query_with_award', methods=['POST'])
def query_participant_with_award():
    data = request.json
    name = data.get('name', '').strip()
    region = data.get('region', '').strip()
    phone = data.get('phone', '').strip()
    organization = data.get('organization', '').strip()
    cert_type = data.get('cert_type', '').strip()

    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('''
        SELECT name, region, phone, organization, certificate_number, award_level FROM participants
        WHERE name = ? AND region = ? AND phone = ? AND organization = ? AND cert_type = ?
    ''', (name, region, phone, organization, cert_type))
    row = c.fetchone()
    conn.close()

    if row:
        return jsonify({
            'found': True,
            'participant': {
                '姓名': row[0],
                '赛区': row[1],
                '手机号': row[2],
                '所在单位': row[3],
                '证书编号': row[4] if row[4] else '',
                '奖项等级': row[5] if row[5] else ''
            }
        })
    else:
        return jsonify({'found': False})


# ==================== 模板管理接口 ====================

@app.route('/api/templates', methods=['GET'])
def get_templates():
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('SELECT cert_type, region, award_level, filename, pdf_data FROM templates')
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
        award_level = row[2]

        if cert_type == 'preliminary':
            if region not in templates[cert_type]:
                templates[cert_type][region] = {}
            templates[cert_type][region][award_level] = {
                'filename': row[3],
                'pdf_data': row[4]
            }
        elif cert_type == 'final':
            if region not in templates[cert_type]:
                templates[cert_type][region] = {}
            templates[cert_type][region][award_level] = {
                'filename': row[3],
                'pdf_data': row[4]
            }
        else:
            templates[cert_type][region] = {
                'filename': row[3],
                'pdf_data': row[4]
            }

    return jsonify(templates)


@app.route('/api/templates/upload', methods=['POST'])
def upload_template():
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    data = request.json
    cert_type = data.get('cert_type')
    region = data.get('region')
    award_level = data.get('award_level', '')
    filename = data.get('filename')
    pdf_data = data.get('pdf_data')

    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    c.execute('DELETE FROM templates WHERE cert_type = ? AND region = ? AND award_level = ?',
              (cert_type, region, award_level))

    c.execute('''
        INSERT INTO templates (cert_type, region, award_level, filename, pdf_data)
        VALUES (?, ?, ?, ?, ?)
    ''', (cert_type, region, award_level, filename, pdf_data))

    conn.commit()
    conn.close()

    return jsonify({'success': True})


@app.route('/api/templates/batch_upload', methods=['POST'])
def batch_upload_templates():
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    data = request.json
    templates_list = data.get('templates', [])

    if not templates_list:
        return jsonify({'success': False, 'error': '没有模板数据'}), 400

    conn = sqlite3.connect('data.db')
    c = conn.cursor()

    success_count = 0
    fail_count = 0
    errors = []

    for template in templates_list:
        try:
            cert_type = template.get('cert_type')
            region = template.get('region')
            award_level = template.get('award_level', '')
            filename = template.get('filename')
            pdf_data = template.get('pdf_data')

            if not cert_type or not region or not filename or not pdf_data:
                fail_count += 1
                errors.append(f"缺少必要字段: {filename}")
                continue

            c.execute('DELETE FROM templates WHERE cert_type = ? AND region = ? AND award_level = ?',
                      (cert_type, region, award_level))

            c.execute('''
                INSERT INTO templates (cert_type, region, award_level, filename, pdf_data)
                VALUES (?, ?, ?, ?, ?)
            ''', (cert_type, region, award_level, filename, pdf_data))
            success_count += 1

        except Exception as e:
            fail_count += 1
            errors.append(f"{template.get('filename', 'unknown')}: {str(e)}")

    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'success_count': success_count,
        'fail_count': fail_count,
        'errors': errors
    })


@app.route('/api/templates/delete', methods=['DELETE'])
def delete_template():
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    data = request.json
    cert_type = data.get('cert_type')
    region = data.get('region')
    award_level = data.get('award_level', '')

    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('DELETE FROM templates WHERE cert_type = ? AND region = ? AND award_level = ?',
              (cert_type, region, award_level))
    conn.commit()
    conn.close()

    return jsonify({'success': True})


@app.route('/api/templates/clear', methods=['DELETE'])
def clear_templates():
    key = request.headers.get('X-Admin-Key')
    if key != ADMIN_KEY:
        return jsonify({'error': '无权操作'}), 401

    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('DELETE FROM templates')
    conn.commit()
    conn.close()

    return jsonify({'success': True})


# ==================== 页面路由 ====================

@app.route('/')
def index():
    return send_from_directory('.', 'query.html')


@app.route('/admin')
def admin():
    return send_from_directory('.', 'admin.html')


@app.route('/<path:path>')
def serve_static(path):
    return send_from_directory('.', path)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)