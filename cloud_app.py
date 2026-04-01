#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
26年NC台账管理系统 - 云端部署版本
支持Render.com免费托管
"""
import os
import json
from datetime import datetime
from flask import Flask, render_template, jsonify, request, send_file, Response
import gzip
import base64
import io
import hashlib
import numpy as np
import pandas as pd

app = Flask(__name__)

# Gzip压缩 - 只压缩文本内容
@app.after_request
def compress_response(response):
    accept_encoding = request.headers.get('Accept-Encoding', '').lower()
    if 'gzip' not in accept_encoding or response.status_code != 200:
        return response
    # 排除二进制文件（Excel、图片等）
    content_type = response.content_type or ''
    if any(x in content_type for x in ['octet-stream', 'excel', 'spreadsheet', 'image', 'pdf']):
        return response
    if 'attachment' in response.headers.get('Content-Disposition', ''):
        return response
    try:
        content = response.get_data()
        gzip_response = Response(gzip.compress(content), content_type=content_type)
        gzip_response.headers['Content-Encoding'] = 'gzip'
        gzip_response.headers['Vary'] = 'Accept-Encoding'
        return gzip_response
    except:
        return response

# Render.com环境变量或本地文件
EXCEL_FILE = os.environ.get('EXCEL_FILE', os.path.expanduser('~/Desktop/26年NC台账勿删.xlsx'))
DATA_FILE = '/tmp/nc_account_data.json'
EMBEDDED_DATA_FILE = os.path.join(os.path.dirname(__file__), 'embedded_data.json')

def load_data():
    """加载数据 - 优先临时文件（最新修改），其次嵌入文件"""
    # 1. 优先从临时文件加载（运行时修改的数据，最新）
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if data and len(data) > 0:
                    return data
        except:
            pass
    
    # 2. 其次从嵌入的数据文件加载（GitHub同步的数据）
    if os.path.exists(EMBEDDED_DATA_FILE):
        try:
            with open(EMBEDDED_DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if data and len(data) > 0:
                    # 复制到临时文件
                    save_data(data)
                    return data
        except:
            pass
    
    return []



# ==================== 操作日志功能 ====================

LOG_FILE = '/tmp/operation_logs.json'
MAX_LOGS = 500  # 最多保留500条日志

def add_log(action, detail, user='system'):
    """添加操作日志"""
    try:
        logs = []
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, 'r') as f:
                logs = json.load(f)
        
        log_entry = {
            'time': datetime.now().isoformat(),
            'action': action,
            'detail': detail,
            'user': user,
            'ip': request.remote_addr if request else 'local'
        }
        
        logs.insert(0, log_entry)
        
        # 限制日志数量
        if len(logs) > MAX_LOGS:
            logs = logs[:MAX_LOGS]
        
        with open(LOG_FILE, 'w') as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f'日志记录失败: {e}')

def get_logs(limit=100):
    """获取操作日志"""
    try:
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, 'r') as f:
                logs = json.load(f)
                return logs[:limit]
    except:
        pass
    return []

@app.route('/api/logs')
def api_logs():
    """获取操作日志API"""
    limit = request.args.get('limit', 100, type=int)
    logs = get_logs(limit)
    return jsonify({'success': True, 'logs': logs})

@app.route('/api/logs/export')
def export_logs():
    """导出操作日志"""
    try:
        logs = get_logs(MAX_LOGS)
        # 生成CSV内容
        csv_content = '时间,操作,详情,用户,IP\n'
        for log in logs:
            time_str = log.get('time', '').replace('T', ' ').split('.')[0]
            action = log.get('action', '')
            detail = log.get('detail', '').replace('"', '""')
            user = log.get('user', 'system')
            ip = log.get('ip', '')
            csv_content += f'"{time_str}","{action}","{detail}","{user}","{ip}"\n'
        
        # 生成文件名
        filename = f'操作日志_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
        return send_file(
            io.BytesIO(csv_content.encode('utf-8-sig')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

def save_data(data):
    """保存数据到临时文件"""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass
    # 同时尝试保存到嵌入文件（本地环境）
    try:
        with open(EMBEDDED_DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass

def sync_to_github(data):
    """同步数据到GitHub（需要配置GITHUB_TOKEN环境变量）"""
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

@app.route('/health')
def health():
    return 'OK', 200

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
        add_log('更新数据', f'索引{index}, 包裹号: {data[index].get("包裹号", "")}')
        save_data(data)
        sync_to_github(data)
        return jsonify({'success': True})
    return jsonify({'success': False}), 404

@app.route('/api/data/<int:index>', methods=['DELETE'])
def delete_data(index):
    data = load_data()
    if 0 <= index < len(data):
        pkg = data[index].get('包裹号', '')
        data.pop(index)
        add_log('删除数据', f'包裹号: {pkg}')
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



@app.route('/api/import-preview', methods=['POST'])
def import_preview():
    """预览导入数据"""
    try:
        data = request.get_json()
        file_data = data.get('fileData', '')
        file_name = data.get('fileName', '')
        
        if not file_data:
            return jsonify({'success': False, 'error': '没有文件数据'}), 400
        
        # 解码Base64数据
        if ',' in file_data:
            file_data = file_data.split(',')[1]
        file_bytes = base64.b64decode(file_data)
        
        # 读取Excel
        df = pd.read_excel(io.BytesIO(file_bytes))
        df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
        df = df.fillna('')
        df = df[df['包裹号'].notna() & (df['包裹号'] != '')]
        
        # 转换数据
        for col in df.columns:
            df[col] = df[col].apply(lambda x: '' if pd.isna(x) or str(x) == 'nan' else str(x))
        
        new_data = df.to_dict('records')
        
        # 获取现有数据
        existing_data = load_data()
        existing_keys = set()
        for item in existing_data:
            key = (str(item.get('日期', '')).strip(), str(item.get('包裹号', '')).strip())
            existing_keys.add(key)
        
        # 计算新增和替换数量
        added = 0
        replaced = 0
        for item in new_data:
            key = (str(item.get('日期', '')).strip(), str(item.get('包裹号', '')).strip())
            if key in existing_keys:
                replaced += 1
            else:
                added += 1
        
        return jsonify({
            'success': True,
            'total': len(new_data),
            'added': added,
            'replaced': replaced,
            'preview': new_data[:5]
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/import-confirm', methods=['POST'])
def import_confirm():
    """确认导入数据"""
    try:
        data = request.get_json()
        file_data = data.get('fileData', '')
        
        if not file_data:
            return jsonify({'success': False, 'error': '没有文件数据'}), 400
        
        # 解码Base64数据
        if ',' in file_data:
            file_data = file_data.split(',')[1]
        file_bytes = base64.b64decode(file_data)
        
        # 读取Excel
        df = pd.read_excel(io.BytesIO(file_bytes))
        df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
        df = df.fillna('')
        df = df[df['包裹号'].notna() & (df['包裹号'] != '')]
        
        for col in df.columns:
            df[col] = df[col].apply(lambda x: '' if pd.isna(x) or str(x) == 'nan' else str(x))
        
        new_data = df.to_dict('records')
        
        # 合并数据
        existing_data = load_data()
        
        def count_filled_fields(item):
            return sum(1 for v in item.values() if str(v).strip() not in ['', 'nan', 'None'])
        
        def pick_more_detailed(item_a, item_b):
            if count_filled_fields(item_b) >= count_filled_fields(item_a):
                return item_b
            return item_a
        
        merged = {}
        for item in existing_data:
            key = (str(item.get('日期', '')).strip(), str(item.get('包裹号', '')).strip())
            if key[1]:
                merged[key] = item
        
        replaced_count = 0
        added_count = 0
        for item in new_data:
            key = (str(item.get('日期', '')).strip(), str(item.get('包裹号', '')).strip())
            if key[1]:
                if key in merged:
                    merged[key] = pick_more_detailed(merged[key], item)
                    replaced_count += 1
                else:
                    merged[key] = item
                    added_count += 1
        
        final_data = list(merged.values())
        final_data.sort(key=lambda x: x.get('日期', ''), reverse=True)
        
        save_data(final_data)
        sync_to_github(final_data)
        add_log('导入数据', f'共{len(final_data)}条, 新增{added_count}条, 替换{replaced_count}条')
        
        return jsonify({
            'success': True,
            'total': len(final_data),
            'added': added_count,
            'replaced': replaced_count
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

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
        
        # 转换为列表前，确保所有值都是基本类型
        for col in df.columns:
            df[col] = df[col].apply(lambda x: '' if pd.isna(x) or str(x) == 'nan' else str(x))
        
        new_data = df.to_dict('records')
        
        def count_filled_fields(item):
            """计算一条记录中非空字段的数量（用于判断哪条更详细）"""
            return sum(1 for v in item.values() if str(v).strip() not in ['', 'nan', 'None'])
        
        def pick_more_detailed(item_a, item_b):
            """比较两条记录，返回更详细的那条"""
            score_a = count_filled_fields(item_a)
            score_b = count_filled_fields(item_b)
            if score_b >= score_a:
                return item_b  # 新数据更详细或相同，用新数据
            return item_a  # 旧数据更详细，保留旧数据
        
        # 获取现有数据
        existing_data = load_data()
        
        # 根据日期+包裹号判断重复，重复则保留更详细的那条
        merged = {}
        
        # 先添加现有数据
        for item in existing_data:
            key = (str(item.get('日期', '')).strip(), str(item.get('包裹号', '')).strip())
            if key[1]:  # 只添加有包裹号的数据
                merged[key] = item
        
        # 再处理新数据
        replaced_count = 0
        added_count = 0
        for item in new_data:
            key = (str(item.get('日期', '')).strip(), str(item.get('包裹号', '')).strip())
            if key[1]:  # 只处理有包裹号的数据
                if key in merged:
                    # 保留更详细的那条
                    merged[key] = pick_more_detailed(merged[key], item)
                    replaced_count += 1
                else:
                    merged[key] = item
                    added_count += 1
        
        # 转换回列表
        final_data = list(merged.values())
        
        # 按日期排序
        final_data.sort(key=lambda x: x.get('日期', ''), reverse=True)
        
        # 保存
        save_data(final_data)
        sync_to_github(final_data)
        add_log('导入数据', f'共{len(final_data)}条, 新增{added_count}条, 替换{replaced_count}条')
        
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

@app.route('/api/calculate-shared', methods=['POST'])
def calculate_shared_expense():
    """计算公摊金额"""
    import traceback
    print("=== 开始计算公摊 ===")
    try:
        # 处理 FormData 格式
        schedule_file = request.files.get('schedule')
        start_date = request.form.get('start_date', '')
        end_date = request.form.get('end_date', '')
        keywords = request.form.get('keywords', '破损,买赔,赔')
        exclude_responsibility = request.form.get('exclude_resp', '')
        
        if not start_date or not end_date:
            return jsonify({'success': False, 'error': '请设置起止日期'}), 400
        
        # 读取排班数据
        if not schedule_file:
            return jsonify({'success': False, 'error': '请上传排班文件'}), 400
        
        schedule_df = pd.read_excel(schedule_file)
        schedule_df['日期'] = pd.to_datetime(schedule_df['日期']).dt.strftime('%Y-%m-%d')
        
        # 剔除非全日制合同工
        if '用工性质' in schedule_df.columns:
            before_count = len(schedule_df)
            schedule_df = schedule_df[schedule_df['用工性质'] != '非全日制劳动合同工']
            excluded_workers = before_count - len(schedule_df)
            if excluded_workers > 0:
                print(f'已剔除 {excluded_workers} 条非全日制合同工记录')
        
        # 自动识别列名（兼容不同格式）
        # 姓名列
        name_col = None
        for c in ['姓名', '人员', '处理人', '名字', 'name']:
            if c in schedule_df.columns:
                name_col = c
                break
        # 班次列
        shift_col_name = None
        for c in ['班次名称', '班次', '班型', '排班', 'shift']:
            if c in schedule_df.columns:
                shift_col_name = c
                break
        # 打卡时间列（判断是否实际出勤）
        clock_col_name = None
        for c in ['实际上班时间', '上班打卡', '打卡时间', '签到时间']:
            if c in schedule_df.columns:
                clock_col_name = c
                break
        
        if not name_col:
            return jsonify({'success': False, 'error': f'排班文件缺少姓名列，当前列名: {list(schedule_df.columns)}'}), 400
        if not shift_col_name:
            return jsonify({'success': False, 'error': f'排班文件缺少班次列，当前列名: {list(schedule_df.columns)}'}), 400
        
        # 班次识别函数
        def classify_shift(shift_str, clock_time=None):
            """识别班次类型: 返回 'day'(白班) 或 'night'(夜班) 或 'rest'(休息) 或 None(未知)"""
            # 先检查是否休息
            if shift_str:
                s = str(shift_str).strip()
                if '休息' in s or '休' == s:
                    return 'rest'
            
            # 检查是否打卡（无打卡时间视为休息/未上班）
            if clock_time is not None:
                ct = str(clock_time).strip()
                if ct in ['', 'nan', 'None', 'null']:
                    return 'rest'  # 没打卡视为休息
            
            if not shift_str or str(shift_str).strip() in ['', 'nan']:
                return None
            s = str(shift_str).strip()
            # 白班: 早班、中班1次、中班3次、中班4次
            if '早班' in s:
                return 'day'
            if '晚班' in s:
                return 'night'
            if '中班' in s:
                import re
                nums = re.findall(r'中班(\d+)次', s)
                if nums:
                    n = int(nums[0])
                    if n in [1, 3, 4]:
                        return 'day'
                    elif n in [2, 5]:
                        return 'night'
            return None  # 无法识别

        # 构建排班字典: {日期: {'day': [白班人员], 'night': [夜班人员], 'all': [全部人员]}}
        schedule = {}
        
        # 使用打卡文件的实际列名
        name_col = name_col  # 已在前面设置
        shift_col = shift_col_name  # 已在前面设置
        clock_col = clock_col_name  # 已在前面设置
        
        for _, row in schedule_df.iterrows():
            date = str(row.get('日期', '')).strip()
            person = str(row.get(name_col, '')).strip() if name_col else ''
            if not date or not person or person in ['', 'nan', 'None']:
                continue
            if date not in schedule:
                schedule[date] = {'day': [], 'night': [], 'all': []}

            # 获取打卡时间（判断是否实际出勤）
            clock_time = row.get(clock_col) if clock_col else None
            
            # 识别班次
            shift_val = str(row.get(shift_col, '')).strip() if shift_col else ''
            shift_type = classify_shift(shift_val, clock_time)

            # 跳过休息的人
            if shift_type == 'rest':
                continue

            # 同一天同一人出现中班+早班，只识别为白班
            if person not in schedule[date]['all']:
                schedule[date]['all'].append(person)

            if shift_type == 'day':
                # 如果之前已在夜班，移到白班（早班优先）
                if person in schedule[date]['night']:
                    schedule[date]['night'].remove(person)
                if person not in schedule[date]['day']:
                    schedule[date]['day'].append(person)
            elif shift_type == 'night':
                # 只有当天没有白班记录时才加入夜班
                if person not in schedule[date]['day'] and person not in schedule[date]['night']:
                    schedule[date]['night'].append(person)
            else:
                # 无法识别班次，归入全部
                pass
        
        # 自动识别单责任人（具体人名）- 包含以下特征的视为单责：
        # 1. 责任方只包含一个具体人名（如：张景莉、吴光辉）
        # 2. 不包含"共责"、"NC"、"验货"、"卸车"等共同责任关键词
        # 从排班文件中提取所有真实人名
        all_persons_in_schedule = set()
        for date_info in schedule.values():
            for p in date_info.get('day', []):
                all_persons_in_schedule.add(p)
            for p in date_info.get('night', []):
                all_persons_in_schedule.add(p)
        
        def is_single_responsibility(resp):
            """判断是否为单责任人（排班文件中的真实人名），返回True表示单责需排除"""
            if not resp or resp == '':
                return False
            # 如果责任方就是排班文件中的某个真实人名，则视为单责剔除
            if resp in all_persons_in_schedule:
                return True
            # 否则不剔除（未拦截、NC、共责等都参与公摊）
            return False
        
        # 筛选时间范围内的破损买赔数据
        nc_data = load_data()
        damaged_items = []
        excluded_items = []  # 被剔除的单责记录
        keyword_list = [k.strip() for k in keywords.split(',') if k.strip()]
        
        print(f"台账数据: {len(nc_data)} 条")
        print(f"关键词: {keyword_list}")
        
        for item in nc_data:
            item_date = str(item.get('日期', '')).strip()
            exception_type = str(item.get('异常情况', ''))
            responsibility = str(item.get('责任方', '')).strip()
            
            # 识别破损买赔相关单子
            is_damaged = any(keyword in exception_type for keyword in keyword_list) if keyword_list else True
            is_in_range = start_date <= item_date <= end_date
            has_amount = item.get('金额') is not None and item.get('金额') != ''
            
            if is_damaged and is_in_range and has_amount:
                amount = float(item.get('金额', 0) or 0)
                if amount > 0:
                    damaged_items.append(item)
                # 自动识别单责任人
                if is_single_responsibility(responsibility):
                    excluded_items.append({
                        'date': item_date,
                        'package': item.get('包裹号', ''),
                        'responsibility': responsibility,
                        'amount': float(item.get('金额', 0) or 0)
                    })
                else:
                    damaged_items.append(item)
        
        # 按日期分组计算
        results = {}
        daily_details = {}
        
        print(f"破损条目: {len(damaged_items)} 条")
        print(f"排除条目: {len(excluded_items)} 条")
        
        # 按日期和班次统计破损金额（白班和夜班分开计算！）
        daily_damage = {}  # {date: {'day': 金额, 'night': 金额}}
        for item in damaged_items:
            date = item.get('日期', '')
            shift = str(item.get('班次', '')).strip()
            amount = float(item.get('金额', 0) or 0)
            
            if date not in daily_damage:
                daily_damage[date] = {'day': 0, 'night': 0}
            
            # 判断班次
            if '白' in shift:
                daily_damage[date]['day'] += amount
            elif '夜' in shift:
                daily_damage[date]['night'] += amount
            else:
                # 无法识别班次，按白班处理
                daily_damage[date]['day'] += amount
        
        # 计算每天每人公摊（白班和夜班分开）
        for date in sorted(daily_damage.keys()):
            day_damage = daily_damage[date]
            day_info = schedule.get(date, {})
            
            # 获取白班和夜班人员
            if isinstance(day_info, list):
                day_people = []
                night_people = []
                all_people = day_info
            else:
                day_people = day_info.get('day', [])
                night_people = day_info.get('night', [])
                all_people = day_info.get('all', [])
            
            # 白班公摊
            if day_people and day_damage['day'] > 0:
                per_person = day_damage['day'] / len(day_people)
                daily_details[date + '_白班'] = {
                    'total': round(day_damage['day'], 2),
                    'people': len(day_people),
                    'per_person': round(per_person, 2),
                    'person_list': day_people,
                    'shift_label': '白班',
                    'day_persons': day_people,
                    'night_persons': []
                }
                for person in day_people:
                    if person not in results:
                        results[person] = {'total': 0, 'dates': []}
                    results[person]['total'] = round(results[person]['total'] + per_person, 2)
                    results[person]['dates'].append({
                        'date': date,
                        'shift': '白班',
                        'amount': round(per_person, 2)
                    })
            
            # 夜班公摊
            if night_people and day_damage['night'] > 0:
                per_person = day_damage['night'] / len(night_people)
                daily_details[date + '_夜班'] = {
                    'total': round(day_damage['night'], 2),
                    'people': len(night_people),
                    'per_person': round(per_person, 2),
                    'person_list': night_people,
                    'shift_label': '夜班',
                    'day_persons': [],
                    'night_persons': night_people
                }
                for person in night_people:
                    if person not in results:
                        results[person] = {'total': 0, 'dates': []}
                    results[person]['total'] = round(results[person]['total'] + per_person, 2)
                    results[person]['dates'].append({
                        'date': date,
                        'shift': '夜班',
                        'amount': round(per_person, 2)
                    })
        
        # 汇总排序
        summary = [{'person': p, 'total': d['total'], 'details': d['dates']} 
                   for p, d in results.items()]
        summary.sort(key=lambda x: x['total'], reverse=True)
        
        grand_total = sum(r['total'] for r in summary)
        
        return jsonify({
            'success': True,
            'start_date': start_date,
            'end_date': end_date,
            'total_damaged': round(sum(daily_damage.values()), 2),
            'days_count': len(daily_damage),
            'people_count': len(results),
            'grand_total': round(grand_total, 2),
            'summary': summary,
            'daily_details': json.loads(json.dumps({k: {kk: (str(vv) if isinstance(vv, (np.integer, np.floating)) else list(vv) if isinstance(vv, (set, frozenset)) else (float(vv) if isinstance(vv, (int, float)) and not isinstance(vv, bool) else vv)) for kk, vv in v.items()} for k, v in daily_details.items()}, default=str)),
            'excluded_count': len(excluded_items),
            'excluded_list': excluded_items
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/export-shared', methods=['POST'])
def export_shared_expense():
    """导出公摊计算结果为Excel"""
    try:
        # 接收计算结果数据
        data = request.json or {}
        summary = data.get('summary', [])
        daily_details = data.get('daily_details', {})
        excluded_list = data.get('excluded_list', [])
        start_date = data.get('start_date', '')
        end_date = data.get('end_date', '')
        
        # 创建结果DataFrame
        rows = []
        for item in summary:
            dates_str = ', '.join([d['date'] + '(' + str(d['amount']) + ')' for d in item['details']])
            rows.append({
                '姓名': item['person'],
                '总公摊金额': item['total'],
                '涉及天数': len(item['details']),
                '明细': dates_str
            })
        
        result_df = pd.DataFrame(rows)
        
        # 创建白班明细和夜班明细（按人名分组）
        # 先按班次分开处理
        day_details = {k: v for k, v in daily_details.items() if '白班' in k}
        night_details = {k: v for k, v in daily_details.items() if '夜班' in k}
        
        # 白班明细：按人名分组
        day_person_details = {}  # {人名: {日期: 金额}}
        all_day_dates = sorted(day_details.keys())
        for key, info in day_details.items():
            date = key.replace('_白班', '')
            day_persons = info.get('day_persons', [])
            per_person = info['per_person']
            
            for person in day_persons:
                if person not in day_person_details:
                    day_person_details[person] = {'总公摊': 0, '天数': 0}
                day_person_details[person][date] = per_person
                day_person_details[person]['总公摊'] += per_person
                day_person_details[person]['天数'] += 1
        
        # 构建白班DataFrame
        day_rows = []
        for person, details in sorted(day_person_details.items(), key=lambda x: -x[1]['总公摊']):
            row = {'姓名': person, '总公摊': round(details['总公摊'], 2), '天数': details['天数']}
            for date in all_day_dates:
                date_key = date.replace('_白班', '')
                if date_key in details:
                    row[date_key] = details[date_key]
            day_rows.append(row)
        day_df = pd.DataFrame(day_rows)
        
        # 夜班明细：按人名分组
        night_person_details = {}
        all_night_dates = sorted(night_details.keys())
        for key, info in night_details.items():
            date = key.replace('_夜班', '')
            night_persons = info.get('night_persons', [])
            per_person = info['per_person']
            
            for person in night_persons:
                if person not in night_person_details:
                    night_person_details[person] = {'总公摊': 0, '天数': 0}
                night_person_details[person][date] = per_person
                night_person_details[person]['总公摊'] += per_person
                night_person_details[person]['天数'] += 1
        
        # 构建夜班DataFrame
        night_rows = []
        for person, details in sorted(night_person_details.items(), key=lambda x: -x[1]['总公摊']):
            row = {'姓名': person, '总公摊': round(details['总公摊'], 2), '天数': details['天数']}
            for date in all_night_dates:
                date_key = date.replace('_夜班', '')
                if date_key in details:
                    row[date_key] = details[date_key]
            night_rows.append(row)
        night_df = pd.DataFrame(night_rows)
        
        # 被剔除的记录
        excluded_df = pd.DataFrame(excluded_list) if excluded_list else pd.DataFrame()
        
        # 写入Excel（多个Sheet）
        excel_path = '/tmp/nc_shared_expense.xlsx'
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            if not result_df.empty:
                result_df.to_excel(writer, index=False, sheet_name='公摊汇总')
            if not day_df.empty:
                day_df.to_excel(writer, index=False, sheet_name='白班明细')
            if not night_df.empty:
                night_df.to_excel(writer, index=False, sheet_name='夜班明细')
            if not excluded_df.empty:
                excluded_df.to_excel(writer, index=False, sheet_name='已剔除记录')
        
        return send_file(excel_path, as_attachment=True, download_name='公摊计算结果.xlsx')
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# Render.com 需要


@app.route('/api/export-template', methods=['POST'])
def export_template():
    """按模板导出数据"""
    try:
        req = request.get_json()
        export_type = req.get('type', 'all')
        data = req.get('data', [])
        
        if not data:
            return jsonify({'success': False, 'error': '没有数据'}), 400
        
        df = pd.DataFrame(data)
        
        # 根据导出类型设置列名
        if export_type == 'by-resp':
            # 责任方汇总
            columns = ['责任方', '数量', '金额']
            df = df[columns] if all(c in df.columns for c in columns) else df
        else:
            # 其他类型保持原列
            preferred_columns = ['日期', '班次', '包裹号', '商品详情', '异常情况', '金额', '责任方', '处理方式', '路由', '处理人', '回款情况']
            existing_cols = [c for c in preferred_columns if c in df.columns]
            df = df[existing_cols] if existing_cols else df
        
        # 创建Excel
        excel_path = '/tmp/nc_export.xlsx'
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='数据')
            
            # 获取工作表并设置列宽
            ws = writer.sheets['数据']
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column].width = adjusted_width
        
        return send_file(excel_path, as_attachment=True, download_name=f'{export_type}_{datetime.now().strftime("%Y%m%d")}.xlsx')
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500




# ==================== 定时同步功能 ====================

# 存储上次同步时间
LAST_SYNC_FILE = '/tmp/last_sync_time.json'

def get_last_sync_time():
    """获取上次同步时间"""
    try:
        if os.path.exists(LAST_SYNC_FILE):
            with open(LAST_SYNC_FILE, 'r') as f:
                return json.load(f).get('last_sync', '')
    except:
        pass
    return ''

def save_last_sync_time():
    """保存同步时间"""
    try:
        with open(LAST_SYNC_FILE, 'w') as f:
            json.dump({'last_sync': datetime.now().isoformat()}, f)
    except:
        pass

@app.route('/api/sync-status')
def sync_status():
    """获取同步状态"""
    return jsonify({
        'last_sync': get_last_sync_time(),
        'data_count': len(load_data()),
        'server_time': datetime.now().isoformat()
    })

@app.route('/api/manual-sync', methods=['POST'])
def manual_sync():
    """手动触发同步"""
    token = os.environ.get('GITHUB_TOKEN')
    if not token:
        return jsonify({'success': False, 'error': '未配置GITHUB_TOKEN，无法同步到GitHub'}), 500
    
    try:
        data = load_data()
        success = sync_to_github(data)
        if success:
            save_last_sync_time()
            return jsonify({'success': True, 'message': '同步成功', 'data_count': len(data)})
        else:
            return jsonify({'success': False, 'error': '同步失败，请检查GITHUB_TOKEN权限'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/schedule-sync', methods=['POST'])
def schedule_sync():
    """设置定时同步（需要外部cron或Render的cron job调用）"""
    try:
        data = load_data()
        sync_to_github(data)
        save_last_sync_time()
        return jsonify({'success': True, 'synced_at': datetime.now().isoformat()})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500




# ==================== 自动备份功能 ====================

BACKUP_DIR = '/tmp/backups'
MAX_BACKUPS = 30  # 保留最近30天备份

def ensure_backup_dir():
    """确保备份目录存在"""
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)

def create_backup():
    """创建数据备份"""
    try:
        ensure_backup_dir()
        
        data = load_data()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = os.path.join(BACKUP_DIR, f'backup_{timestamp}.json')
        
        with open(backup_file, 'w', encoding='utf-8') as f:
            json.dump({
                'timestamp': datetime.now().isoformat(),
                'data_count': len(data),
                'data': data
            }, f, ensure_ascii=False, indent=2)
        
        # 清理旧备份（保留最近30天）
        clean_old_backups()
        
        add_log('自动备份', f'备份成功: {len(data)}条数据')
        return True
    except Exception as e:
        print(f'备份失败: {e}')
        return False

def clean_old_backups():
    """清理旧备份"""
    try:
        ensure_backup_dir()
        files = sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith('backup_')])
        
        while len(files) > MAX_BACKUPS:
            old_file = os.path.join(BACKUP_DIR, files[0])
            os.remove(old_file)
            files.pop(0)
    except Exception as e:
        print(f'清理旧备份失败: {e}')

def get_backup_list():
    """获取备份列表"""
    try:
        ensure_backup_dir()
        files = sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith('backup_')], reverse=True)
        
        backups = []
        for f in files:
            filepath = os.path.join(BACKUP_DIR, f)
            stat = os.stat(filepath)
            backups.append({
                'filename': f,
                'size': stat.st_size,
                'time': datetime.fromtimestamp(stat.st_mtime).isoformat()
            })
        return backups
    except:
        return []

@app.route('/api/backups')
def api_backups():
    """获取备份列表"""
    backups = get_backup_list()
    return jsonify({'success': True, 'backups': backups})

@app.route('/api/backup-now', methods=['POST'])
def api_backup_now():
    """立即创建备份"""
    success = create_backup()
    if success:
        return jsonify({'success': True, 'message': '备份成功'})
    return jsonify({'success': False, 'error': '备份失败'}), 500

@app.route('/api/restore-backup/<filename>', methods=['POST'])
def api_restore_backup(filename):
    """恢复备份"""
    try:
        filepath = os.path.join(BACKUP_DIR, filename)
        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '备份文件不存在'}), 404
        
        with open(filepath, 'r') as f:
            backup_data = json.load(f)
        
        data = backup_data.get('data', [])
        save_data(data)
        sync_to_github(data)
        add_log('恢复备份', f'从 {filename} 恢复 {len(data)} 条数据')
        
        return jsonify({'success': True, 'message': f'已恢复 {len(data)} 条数据'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/download-backup/<filename>')
def api_download_backup(filename):
    """下载备份文件"""
    try:
        filepath = os.path.join(BACKUP_DIR, filename)
        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'}), 404
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
