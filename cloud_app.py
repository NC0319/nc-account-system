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
        
        # 转换为列表前，确保所有值都是基本类型
        for col in df.columns:
            df[col] = df[col].apply(lambda x: '' if pd.isna(x) or str(x) == 'nan' else str(x))
        
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

@app.route('/api/calculate-shared', methods=['POST'])
def calculate_shared_expense():
    """计算公摊金额"""
    try:
        # 处理 FormData 格式
        schedule_file = request.files.get('schedule')
        start_date = request.form.get('start_date', '')
        end_date = request.form.get('end_date', '')
        keywords = request.form.get('keywords', '破损,买赔,赔')
        exclude_responsibility = request.form.get('exclude_resp', '张三,李四,王五')  # 排除的具体人名
        
        if not schedule_file:
            return jsonify({'success': False, 'error': '请上传排班文件'}), 400
        
        if not start_date or not end_date:
            return jsonify({'success': False, 'error': '请设置起止日期'}), 400
        
        # 读取排班数据
        schedule_df = pd.read_excel(schedule_file)
        schedule_df['日期'] = pd.to_datetime(schedule_df['日期']).dt.strftime('%Y-%m-%d')
        
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
        # 检测排班文件是否有班次列和打卡时间列
        shift_col = None
        clock_col = None
        for col in schedule_df.columns:
            col_str = str(col)
            if '班次' in col_str or '班型' in col_str or '排班' in col_str:
                shift_col = col
            if '打卡' in col_str or '时间' in col_str or 'Clock' in col_str:
                clock_col = col
        
        # 如果没找到打卡列，尝试常见列名
        if not clock_col:
            for col in ['上班打卡', '打卡时间', 'ClockIn', 'clock_in', '签到']:
                if col in schedule_df.columns:
                    clock_col = col
                    break

        for _, row in schedule_df.iterrows():
            date = str(row.get('日期', '')).strip()
            person = str(row.get('处理人', row.get('人员', row.get('姓名', '')))).strip()
            if not date or not person or person in ['', 'nan']:
                continue
            if date not in schedule:
                schedule[date] = {'day': [], 'night': [], 'all': []}

            # 获取打卡时间
            clock_time = row.get(clock_col) if clock_col else None
            
            # 识别班次
            shift_type = None
            if shift_col:
                shift_val = str(row.get(shift_col, '')).strip()
                shift_type = classify_shift(shift_val, clock_time)
            else:
                # 没有班次列，尝试从其他字段识别
                for col in schedule_df.columns:
                    val = str(row.get(col, '')).strip()
                    t = classify_shift(val, clock_time)
                    if t:
                        shift_type = t
                        break

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
        single_resp_keywords = ['共责', 'NC', '验货', '卸车', '发货组', '接货仓', '分拣机', '传送带', '滑槽', '进港', '拨货', '无人', 'T20', '模组', '全仓', '撞', '烂', '挤', '污染', '错分', '被']
        
        def is_single_responsibility(resp):
            """判断是否为单责任人（具体人名），返回True表示单责需排除"""
            if not resp or resp == '':
                return False
            resp_upper = resp.upper()
            # 如果包含"共责"字样，说明是多责，不排除
            if '共责' in resp:
                return False
            # 如果包含共同责任关键词，说明是多责，不排除
            for kw in single_resp_keywords:
                if kw in resp:
                    return False
            # 否则视为单责，需要排除
            return True
        
        # 筛选时间范围内的破损买赔数据
        nc_data = load_data()
        damaged_items = []
        excluded_items = []  # 被剔除的单责记录
        keyword_list = [k.strip() for k in keywords.split(',') if k.strip()]
        
        for item in nc_data:
            item_date = str(item.get('日期', '')).strip()
            exception_type = str(item.get('异常情况', ''))
            responsibility = str(item.get('责任方', '')).strip()
            
            # 识别破损买赔相关单子
            is_damaged = any(keyword in exception_type for keyword in keyword_list) if keyword_list else True
            is_in_range = start_date <= item_date <= end_date
            
            if is_damaged and is_in_range and item.get('金额'):
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
        
        # 先按日期统计破损金额
        daily_damage = {}
        for item in damaged_items:
            date = item.get('日期', '')
            amount = float(item.get('金额', 0) or 0)
            daily_damage[date] = daily_damage.get(date, 0) + amount
        
        # 计算每天每人公摊
        for date, total_damage in daily_damage.items():
            day_info = schedule.get(date, {})
            # 兼容旧格式（列表）和新格式（字典）
            if isinstance(day_info, list):
                people = day_info
                shift_label = ''
            else:
                day_people = day_info.get('day', [])
                night_people = day_info.get('night', [])
                all_people = day_info.get('all', [])
                # 有班次区分时用 all（白班+夜班合并），无班次时用 all
                people = all_people if all_people else []
                # 构建班次标签
                labels = []
                if day_people:
                    labels.append('白班:' + ','.join(day_people))
                if night_people:
                    labels.append('夜班:' + ','.join(night_people))
                shift_label = ' | '.join(labels) if labels else ''

            if people:
                per_person = total_damage / len(people)
                daily_details[date] = {
                    'total': round(total_damage, 2),
                    'people': len(people),
                    'per_person': round(per_person, 2),
                    'person_list': people,
                    'shift_label': shift_label
                }
                for person in people:
                    if person not in results:
                        results[person] = {'total': 0, 'dates': []}
                    results[person]['total'] = round(results[person]['total'] + per_person, 2)
                    results[person]['dates'].append({
                        'date': date,
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
            'daily_details': daily_details,
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
        
        # 创建明细DataFrame
        detail_rows = []
        for date, info in daily_details.items():
            for person in info.get('person_list', []):
                detail_rows.append({
                    '日期': date,
                    '当天破损总额': info['total'],
                    '上班人数': info['people'],
                    '每人公摊': info['per_person'],
                    '上班人员': ', '.join(info.get('person_list', []))
                })
        detail_df = pd.DataFrame(detail_rows)
        
        # 被剔除的记录
        excluded_df = pd.DataFrame(excluded_list) if excluded_list else pd.DataFrame()
        
        # 写入Excel（多个Sheet）
        excel_path = '/tmp/nc_shared_expense.xlsx'
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            if not result_df.empty:
                result_df.to_excel(writer, index=False, sheet_name='公摊汇总')
            if not detail_df.empty:
                detail_df.to_excel(writer, index=False, sheet_name='每日明细')
            if not excluded_df.empty:
                excluded_df.to_excel(writer, index=False, sheet_name='已剔除记录')
        
        return send_file(excel_path, as_attachment=True, download_name='公摊计算结果.xlsx')
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# Render.com 需要
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
