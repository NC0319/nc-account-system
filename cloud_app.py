#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
26年NC台账管理系统 - 云端部署版本
支持Render.com免费托管
"""
import os
import json
from datetime import datetime
from flask import Flask, render_template, jsonify, request, send_file
import pandas as pd

app = Flask(__name__)

# Render.com环境变量或本地文件
EXCEL_FILE = os.environ.get('EXCEL_FILE', os.path.expanduser('~/Desktop/26年NC台账勿删.xlsx'))
DATA_FILE = '/tmp/nc_account_data.json'
EMBEDDED_DATA_FILE = os.path.join(os.path.dirname(__file__), 'embedded_data.json')

def load_data():
    """加载数据 - 优先嵌入文件，其次临时文件"""
    # 1. 从嵌入的数据文件加载（持久化数据）
    if os.path.exists(EMBEDDED_DATA_FILE):
        try:
            with open(EMBEDDED_DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if data and len(data) > 0:
                    return data
        except:
            pass
    
    # 2. 其次从临时文件加载
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if data and len(data) > 0:
                    return data
        except:
            pass
    
    return []

def save_data(data):
    """保存数据到临时文件"""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass

def sync_to_github(data):
    """同步数据到GitHub（需要配置GITHUB_TOKEN环境变量）"""
    import base64
    import hashlib
    
    token = os.environ.get('GITHUB_TOKEN')
    if not token:
        return False
    
    try:
        import requests
        
        # GitHub API 配置
        owner = 'NC0319'
        repo = 'nc-account-system'
        path = 'embedded_data.json'
        
        # 获取当前文件的 SHA
        api_url = f'https://api.github.com/repos/{owner}/{repo}/contents/{path}'
        headers = {
            'Authorization': f'token {token}',
            'Accept': 'application/vnd.github.v3+json'
        }
        
        r = requests.get(api_url, headers=headers)
        sha = r.json().get('sha', '') if r.status_code == 200 else ''
        
        # 准备新内容
        content = json.dumps(data, ensure_ascii=False, indent=2)
        content_b64 = base64.b64encode(content.encode('utf-8')).decode('utf-8')
        
        # 更新文件
        payload = {
            'message': '自动同步数据更新',
            'content': content_b64,
            'sha': sha
        }
        
        r = requests.put(api_url, headers=headers, json=payload)
        return r.status_code in [200, 201]
    except Exception as e:
        print(f"同步GitHub失败: {e}")
        return False

def save_to_excel(data):
    """同步保存到Excel"""
    try:
        df = pd.DataFrame(data)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        # 云端写入到挂载的磁盘，本地写入桌面
        excel_path = os.environ.get('EXCEL_FILE', EXCEL_FILE)
        
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='台账')
        return True
    except Exception as e:
        print(f"保存Excel失败: {e}")
        return False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data', methods=['GET'])
def get_data():
    return jsonify(load_data())

@app.route('/api/data', methods=['POST'])
def add_data():
    data = load_data()
    data.append(request.json)
    save_data(data)
    sync_to_github(data)
    return jsonify({'success': True})

@app.route('/api/data/<int:index>', methods=['PUT'])
def update_data(index):
    data = load_data()
    if 0 <= index < len(data):
        data[index] = request.json
        save_data(data)
        sync_to_github(data)
        return jsonify({'success': True})
    return jsonify({'success': False}), 404

@app.route('/api/data/<int:index>', methods=['DELETE'])
def delete_data(index):
    data = load_data()
    if 0 <= index < len(data):
        data.pop(index)
        save_data(data)
        sync_to_github(data)
        return jsonify({'success': True})
    return jsonify({'success': False}), 404

@app.route('/api/export')
def export_excel():
    data = load_data()
    excel_path = '/tmp/nc_account_export.xlsx'
    df = pd.DataFrame(data)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.to_excel(excel_path, index=False)
    return send_file(excel_path, as_attachment=True)

@app.route('/api/import', methods=['POST'])
def import_excel():
    """导入Excel文件，支持重复数据替换"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '没有上传文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': '文件名为空'}), 400
        
        # 读取上传的Excel
        df = pd.read_excel(file)
        df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
        df = df.fillna('')
        
        # 清理空行（包裹号为空的不导入）
        df = df[df['包裹号'].notna() & (df['包裹号'] != '')]
        df = df.drop_duplicates(subset=['包裹号', '日期'], keep='last')
        
        new_data = df.to_dict('records')
        
        # 获取现有数据
        existing_data = load_data()
        
        # 根据包裹号+日期判断重复，重复则替换，否则追加
        merged = {}
        
        # 先添加现有数据（用包裹号+日期作为key）
        for i, item in enumerate(existing_data):
            key = (str(item.get('包裹号', '')).strip(), str(item.get('日期', '')).strip())
            if key[0]:  # 只添加有包裹号的数据
                merged[key] = item
        
        # 再用新数据覆盖或添加
        replaced_count = 0
        added_count = 0
        for item in new_data:
            key = (str(item.get('包裹号', '')).strip(), str(item.get('日期', '')).strip())
            if key[0]:  # 只处理有包裹号的数据
                if key in merged:
                    replaced_count += 1
                else:
                    added_count += 1
                merged[key] = item
        
        # 转换回列表
        final_data = list(merged.values())
        
        # 按日期排序
        final_data.sort(key=lambda x: x.get('日期', ''), reverse=True)
        
        # 保存
        save_data(final_data)
        sync_to_github(final_data)
        
        return jsonify({
            'success': True,
            'total': len(final_data),
            'added': added_count,
            'replaced': replaced_count
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/batch-paid', methods=['POST'])
def batch_mark_paid():
    """批量标记回款"""
    try:
        indices = json.loads(request.form.get('indices', '[]'))
        data = load_data()
        count = 0
        for idx in indices:
            if 0 <= idx < len(data):
                data[idx]['回款情况'] = '√'
                count += 1
        save_data(data)
        sync_to_github(data)
        return jsonify({'success': True, 'count': count})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# Render.com 需要
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
