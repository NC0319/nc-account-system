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
    """加载数据"""
    # 优先从临时文件加载（运行时数据）
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    # 其次从嵌入的数据文件加载（部署时自带）
    if os.path.exists(EMBEDDED_DATA_FILE):
        with open(EMBEDDED_DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            save_data(data)
            return data
    # 最后尝试从Excel加载
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
            df = df.fillna('')
            data = df.to_dict('records')
            save_data(data)
            return data
        except:
            return []
    return []

def save_data(data):
    """保存数据"""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass

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
    save_to_excel(data)
    return jsonify({'success': True})

@app.route('/api/data/<int:index>', methods=['PUT'])
def update_data(index):
    data = load_data()
    if 0 <= index < len(data):
        data[index] = request.json
        save_data(data)
        save_to_excel(data)
        return jsonify({'success': True})
    return jsonify({'success': False}), 404

@app.route('/api/data/<int:index>', methods=['DELETE'])
def delete_data(index):
    data = load_data()
    if 0 <= index < len(data):
        data.pop(index)
        save_data(data)
        save_to_excel(data)
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

@app.route('/api/data/batch', methods=['POST'])
def batch_import():
    """批量导入数据"""
    try:
        new_data = request.json.get('data', [])
        save_data(new_data)
        save_to_excel(new_data)
        return jsonify({'success': True, 'count': len(new_data)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# Render.com 需要
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
