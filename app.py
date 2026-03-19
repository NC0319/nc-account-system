#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
26年NC台账管理系统 - Flask后端
"""
import os
import json
from datetime import datetime
from flask import Flask, render_template, jsonify, request, send_file
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)

# 配置
EXCEL_FILE = os.path.expanduser('~/Desktop/26年NC台账勿删.xlsx')
DATA_FILE = os.path.join(os.path.dirname(__file__), 'data.json')

def load_excel_data():
    """加载Excel数据"""
    try:
        df = pd.read_excel(EXCEL_FILE)
        # 处理日期格式
        df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
        # 填充NaN为None
        df = df.fillna('')
        # 转换为字典列表
        data = df.to_dict('records')
        return data
    except Exception as e:
        print(f"加载Excel失败: {e}")
        return []

def save_to_excel(data):
    """保存数据到Excel"""
    try:
        df = pd.DataFrame(data)
        # 移除空列
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        # 保存到Excel
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='台账')
        return True
    except Exception as e:
        print(f"保存Excel失败: {e}")
        return False

# 初始化数据
def init_data():
    """初始化数据文件"""
    if not os.path.exists(DATA_FILE):
        data = load_excel_data()
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')

@app.route('/api/data', methods=['GET'])
def get_data():
    """获取所有数据"""
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    else:
        data = load_excel_data()
    return jsonify(data)

@app.route('/api/data', methods=['POST'])
def add_data():
    """添加新数据"""
    new_item = request.json
    
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    else:
        data = []
    
    data.append(new_item)
    
    # 保存到JSON
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    # 同步到Excel
    save_to_excel(data)
    
    return jsonify({'success': True})

@app.route('/api/data/<int:index>', methods=['PUT'])
def update_data(index):
    """更新数据"""
    updated_item = request.json
    
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    else:
        return jsonify({'success': False, 'error': '数据文件不存在'}), 404
    
    if 0 <= index < len(data):
        data[index] = updated_item
        
        # 保存到JSON
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # 同步到Excel
        save_to_excel(data)
        
        return jsonify({'success': True})
    else:
        return jsonify({'success': False, 'error': '索引超出范围'}), 404

@app.route('/api/data/<int:index>', methods=['DELETE'])
def delete_data(index):
    """删除数据"""
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    else:
        return jsonify({'success': False, 'error': '数据文件不存在'}), 404
    
    if 0 <= index < len(data):
        data.pop(index)
        
        # 保存到JSON
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # 同步到Excel
        save_to_excel(data)
        
        return jsonify({'success': True})
    else:
        return jsonify({'success': False, 'error': '索引超出范围'}), 404

@app.route('/api/export')
def export_excel():
    """导出Excel"""
    if os.path.exists(DATA_FILE):
        return send_file(EXCEL_FILE, as_attachment=True)
    else:
        return '数据文件不存在', 404

@app.route('/api/sync', methods=['POST'])
def sync_from_excel():
    """从Excel同步数据"""
    data = load_excel_data()
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return jsonify({'success': True, 'count': len(data)})

if __name__ == '__main__':
    init_data()
    # 获取本机IP
    import socket
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    
    print(f"\n{'='*50}")
    print(f"🎉 26年NC台账管理系统 已启动!")
    print(f"{'='*50}")
    print(f"📍 访问地址:")
    print(f"   本机: http://localhost:5000")
    print(f"   局域网: http://{local_ip}:5000")
    print(f"{'='*50}\n")
    
    app.run(host='0.0.0.0', port=5001, debug=True)
