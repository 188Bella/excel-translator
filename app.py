from flask import Flask, render_template, request, send_file, jsonify
import os
from werkzeug.utils import secure_filename
import requests
import json
import time
import re
import hashlib
import random
import xlwings as xw  # 导入 xlwings

app = Flask(__name__)
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(basedir, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB限制
app.config['TERM_BASE_FILE'] = 'term_base.json'  # 术语库文件

# 确保上传目录存在
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# 在应用启动时确保术语库存在
def ensure_term_base_exists():
    term_base_path = app.config['TERM_BASE_FILE']
    if not os.path.exists(term_base_path):
        try:
            with open(term_base_path, 'w', encoding='utf-8') as f:
                json.dump({}, f, ensure_ascii=False, indent=2)
            print(f"已创建术语库文件: {term_base_path}")
        except Exception as e:
            print(f"创建术语库文件失败: {str(e)}")

# 加载术语库
def load_term_base():
    term_base_path = app.config['TERM_BASE_FILE']
    print(f"尝试加载术语库: {term_base_path}")
    
    # 获取当前工作目录
    current_dir = os.getcwd()
    print(f"当前工作目录: {current_dir}")
    
    # 检查文件是否存在
    file_exists = os.path.exists(term_base_path)
    print(f"术语库文件存在: {file_exists}")
    
    if file_exists:
        try:
            with open(term_base_path, 'r', encoding='utf-8') as f:
                term_base = json.load(f)
                print(f"成功加载术语库，包含 {len(term_base)} 个术语")
                return term_base
        except Exception as e:
            print(f"加载术语库失败: {str(e)}")
            return {}
    else:
        # 如果术语库文件不存在，尝试创建
        print("术语库文件不存在，尝试创建...")
        try:
            with open(term_base_path, 'w', encoding='utf-8') as f:
                json.dump({}, f, ensure_ascii=False, indent=2)
            print(f"已创建空的术语库文件: {term_base_path}")
            return {}
        except Exception as e:
            print(f"创建术语库文件失败: {str(e)}")
            return {}

# 保存术语库
def save_term_base(term_base):
    term_base_path = app.config['TERM_BASE_FILE']
    try:
        with open(term_base_path, 'w', encoding='utf-8') as f:
            json.dump(term_base, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"保存术语库失败: {str(e)}")
        return False

# 简单的语言检测函数
def detect_language(text):
    # 检查是否包含中文字符
    if re.search(r'[\u4e00-\u9fff]', text):
        return 'zh'
    # 检查是否主要是英文字母
    elif re.search(r'[a-zA-Z]', text) and not re.search(r'[\u4e00-\u9fff]', text):
        return 'en'
    else:
        return 'unknown'

# 在术语库中查找翻译
def lookup_term_base(text, from_lang, to_lang, term_base):
    if from_lang == 'zh' and to_lang == 'en':
        # 中文到英文的查找
        return term_base.get(text, None)
    elif from_lang == 'en' and to_lang == 'zh':
        # 英文到中文的查找 - 需要遍历所有值
        for key, value in term_base.items():
            if value.lower() == text.lower():
                return key
    return None

# 调用翻译API
def translate_text(query, from_lang='auto', to_lang='en', term_base=None):
    """
    使用百度翻译API翻译文本，优先使用术语库
    """
    # 如果提供了术语库，先尝试在术语库中查找
    if term_base is not None:
        term_translation = lookup_term_base(query, from_lang, to_lang, term_base)
        if term_translation:
            print(f"使用术语库翻译: {query} -> {term_translation}")
            return term_translation
    
    # 从环境变量获取密钥
    app_id = os.environ.get("BAIDU_APP_ID")
    secret_key = os.environ.get("BAIDU_SECRET_KEY")
    
    # 检查是否成功获取密钥
    if not app_id or not secret_key:
        print("错误: 未找到百度翻译API的APP ID或密钥")
        return f"[翻译错误: 未配置API密钥] {query}"
    
    # 生成随机数（salt）
    salt = str(random.randint(10000, 99999))
    
    # 计算签名（sign）
    sign_str = app_id + query + salt + secret_key
    sign = hashlib.md5(sign_str.encode()).hexdigest()
    
    # 构建请求URL
    url = 'https://fanyi-api.baidu.com/api/trans/vip/translate'
    
    # 构建请求参数
    params = {
        'q': query,
        'from': from_lang,
        'to': to_lang,
        'appid': app_id,
        'salt': salt,
        'sign': sign
    }
    
    try:
        # 发送请求
        response = requests.get(url, params=params, timeout=10)
        
        # 检查响应内容是否为JSON
        if not response.text.strip().startswith('{'):
            print(f"百度翻译API返回非JSON响应: {response.text[:200]}...")
            return f"[翻译错误: 无效的API响应] {query}"
        
        result = response.json()
        
        # 解析并返回翻译结果
        if 'trans_result' in result:
            return result['trans_result'][0]['dst']
        else:
            error_msg = result.get('error_msg', '未知错误')
            error_code = result.get('error_code', '未知错误码')
            print(f"百度翻译API错误: 错误码={error_code}, 错误信息={error_msg}, 完整响应={result}")
            return f"[翻译错误: {error_code}] {query}"
            
    except requests.exceptions.Timeout:
        print("百度翻译API请求超时")
        return f"[翻译超时] {query}"
    except requests.exceptions.RequestException as e:
        print(f"网络请求错误: {str(e)}")
        return f"[网络错误] {query}"
    except json.JSONDecodeError as e:
        print(f"解析百度翻译API响应失败: {str(e)}")
        # 打印响应内容的前200个字符以便调试
        print(f"响应内容: {response.text[:200]}...")
        return f"[解析错误] {query}"

# 检查特殊格式的函数
def is_special_format(text):
    """
    检查文本是否为特殊格式（如URL、电子邮件、数字等）
    如果是，则不需要翻译
    """
    # URL检测
    if re.match(r'https?://\S+', text):
        return True
    
    # 电子邮件检测
    if re.match(r'\S+@\S+\.\S+', text):
        return True
    
    # 纯数字检测
    if re.match(r'^\d+$', text):
        return True
    
    # 产品代码/编号格式检测（数字、字母和特定分隔符的组合）
    # 例如：302*302、12325-D-3221、ABC-123-XYZ
    if re.match(r'^[A-Za-z0-9]+([-_\*\./\\][A-Za-z0-9]+)*$', text):
       return True
    
    # 检查是否同时包含中文和英文（且英文字母超过5个）
    # 这种情况通常表示已经翻译过的内容
    if re.search(r'[\u4e00-\u9fff]', text) and re.search(r'[a-zA-Z]', text):
        # 计算英文字母数量
        english_chars = re.findall(r'[a-zA-Z]', text)
        if len(english_chars) > 5:
            return True
    
    # 日期格式检测（简单版本）
    date_patterns = [
        r'\d{4}-\d{2}-\d{2}',
        r'\d{2}/\d{2}/\d{4}',
        r'\d{4}/\d{2}/\d{2}',
        r'\d{2}-\d{2}-\d{4}'
    ]
    for pattern in date_patterns:
        if re.match(pattern, text):
            return True
    
    return False

# 函数：检查单元格是否已经翻译过
def is_already_translated(text):
    """
    检查文本是否已经是中英文对照格式
    判断标准：包含换行符，且换行符前后文本的语言不同
    """
    # 如果文本中没有换行符，肯定不是翻译后的格式
    if '\n' not in text:
        return False
    
    # 分割文本
    parts = text.split('\n', 1)  # 只分割第一个换行符
    if len(parts) < 2:
        return False
    
    part1, part2 = parts
    
    # 如果两部分都是空字符串，不算翻译过的
    if not part1.strip() or not part2.strip():
        return False
    
    # 检测两部分语言
    lang1 = detect_language(part1)
    lang2 = detect_language(part2)
    
    # 如果两部分语言不同，则认为已经翻译过
    if (lang1 == 'zh' and lang2 == 'en') or (lang1 == 'en' and lang2 == 'zh'):
        return True
    
    return False

# 首页路由
@app.route('/')
def index():
    return render_template('index.html')

# 文件上传和处理路由
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '没有选择文件'})  # 添加success字段
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'})  # 添加success字段
    
    if file and allowed_file(file.filename):
        # 保存上传的文件
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # 处理Excel文件
            output_filename = process_excel(filepath, filename)
            download_url = f"/download/{output_filename}"
            
            return jsonify({
                'success': True,
                'message': '文件处理成功',
                'download_url': download_url
            })
        except Exception as e:
            return jsonify({'success': False, 'error': f'处理文件时出错: {str(e)}'})  # 添加success字段
    
    return jsonify({'success': False, 'error': '不支持的文件类型'})  # 添加success字段

# 下载文件路由
@app.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    print(f"尝试下载文件: {filepath}")
    # 检查文件是否存在
    if not os.path.exists(filepath):
        print(f"文件不存在: {filepath}")
        return jsonify({'success': False, 'error': '文件不存在'}), 404  # 添加success字段
    return send_file(filepath, as_attachment=True)

# 术语库管理路由
@app.route('/term_base', methods=['GET', 'POST', 'DELETE'])
def manage_term_base():
    term_base = load_term_base()
    
    if request.method == 'GET':
        # 获取术语库内容
        return jsonify(term_base)
    
    elif request.method == 'POST':
        # 添加或更新术语
        data = request.get_json()
        if not data or 'term' not in data or 'translation' not in data:
            return jsonify({'error': '缺少术语或翻译'})
        
        term = data['term'].strip()
        translation = data['translation'].strip()
        
        if not term or not translation:
            return jsonify({'error': '术语和翻译不能为空'})
        
        term_base[term] = translation
        if save_term_base(term_base):
            return jsonify({'success': True, 'message': '术语添加成功'})
        else:
            return jsonify({'error': '保存术语库失败'})
    
    elif request.method == 'DELETE':
        # 删除术语
        data = request.get_json()
        if not data or 'term' not in data:
            return jsonify({'error': '缺少术语'})
        
        term = data['term'].strip()
        if term in term_base:
            del term_base[term]
            if save_term_base(term_base):
                return jsonify({'success': True, 'message': '术语删除成功'})
            else:
                return jsonify({'error': '保存术语库失败'})
        else:
            return jsonify({'error': '术语不存在'})

# 检查文件类型
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ['xlsx', 'xls']

# 处理Excel文件的核心函数 - 使用 xlwings
def process_excel(filepath, original_filename):
    # 加载术语库
    term_base = load_term_base()
    
    # 启动 Excel 应用程序（不可见模式）
    app_xw = xw.App(visible=False)
    
    try:
        # 打开工作簿
        wb = app_xw.books.open(filepath)
        
        # 遍历所有工作表
        for sheet in wb.sheets:
            print(f"正在处理工作表: {sheet.name}")
            
            # 获取工作表中所有已使用的单元格范围
            used_range = sheet.used_range
            if not used_range:
                print(f"工作表 {sheet.name} 没有数据，跳过")
                continue
                
            # 获取行数和列数
            row_count = used_range.rows.count
            col_count = used_range.columns.count
            
            print(f"工作表 {sheet.name} 有 {row_count} 行, {col_count} 列数据")
            
            # 遍历所有单元格
            for row in range(1, row_count + 1):
                for col in range(1, col_count + 1):
                    # 获取单元格
                    cell = sheet.range((row, col))
                    if cell.value and isinstance(cell.value, str):
                        # 检查是否为特殊格式
                        if is_special_format(cell.value):
                            print(f"跳过特殊格式单元格: {cell.value}")
                            continue
                        
                        # 检查是否已翻译
                        if is_already_translated(cell.value):
                            print(f"跳过已翻译单元格: {cell.value}")
                            continue
                        
                        # 检测文本语言
                        lang = detect_language(cell.value)
                        
                        # 根据语言决定翻译方向
                        if lang == 'zh':
                            translated = translate_text(cell.value, 'zh', 'en', term_base)
                            cell.value = f"{cell.value}\n{translated}"
                            print(f"翻译中文: {cell.value} -> {translated}")
                        elif lang == 'en':
                            translated = translate_text(cell.value, 'en', 'zh', term_base)
                            cell.value = f"{cell.value}\n{translated}"
                            print(f"翻译英文: {cell.value} -> {translated}")
        
        # 生成输出文件名
        name, ext = os.path.splitext(original_filename)
        if not ext:
            ext = '.xlsx'
        output_filename = f"{name}_translated{ext}"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        print(f"原始文件名: {original_filename}")
        print(f"输出文件名: {output_filename}")
        print(f"输出路径: {output_path}")

        # --- 修改部分开始 ---
        # 保存处理后的工作簿
        wb.save(output_path)
        print("保存命令已发出，等待Excel进程完成...")

        # 添加延迟，等待文件保存完成
        # 这是解决文件保存失败的关键
        import time
        time.sleep(2) # 等待2秒，可以根据文件大小调整这个时间

        # 检查文件是否成功保存
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"文件成功保存: {output_path}, 大小: {file_size} 字节")
            if file_size == 0:
                 print("警告：保存的文件大小为0字节，可能存在保存问题。")
        else:
            print(f"警告：文件保存后检查失败，文件不存在: {output_path}")
            # 尝试再次检查，有时文件系统可能有延迟
            time.sleep(1)
            if os.path.exists(output_path):
                 file_size = os.path.getsize(output_path)
                 print(f"第二次检查成功: {output_path}, 大小: {file_size} 字节")
            else:
                 print(f"第二次检查仍然失败: {output_path}")
                 raise Exception(f"文件保存失败: {output_path}") # 抛出异常，让上层捕获
        # --- 修改部分结束 ---

        # 关闭工作簿
        wb.close()
        
        return output_filename
        
    except Exception as e:
        # 确保在异常情况下也关闭工作簿和应用程序
        if 'wb' in locals():
            try:
                wb.close()
            except:
                pass # 忽略关闭工作簿时的错误
        print(f"处理Excel时出错: {str(e)}")
        raise e # 重新抛出异常
    finally:
        # 确保总是退出 Excel 应用程序
        try:
            app_xw.quit()
        except Exception as e:
            print(f"关闭Excel应用程序时出错: {str(e)}")

# 测试百度翻译API是否可用
def test_baidu_api():
    """测试百度翻译API是否可用"""
    test_text = "你好"
    result = translate_text(test_text, 'zh', 'en')
    
    if result.startswith("[翻译错误") or result.startswith("[网络错误") or result.startswith("[解析错误"):
        print(f"百度翻译API测试失败: {result}")
        return False
    else:
        print(f"百度翻译API测试成功: '{test_text}' -> '{result}'")
        return True

if __name__ == '__main__':
    # 确保术语库文件存在
    ensure_term_base_exists()
    
    # 测试翻译函数（仅在直接运行时执行）
    api_available = test_baidu_api()
    if not api_available:
        print("警告: 百度翻译API可能不可用，请检查API密钥和网络连接")
    
    app.run(debug=True)