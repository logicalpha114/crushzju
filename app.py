from flask import Flask, render_template, request, jsonify, session, redirect, url_for
import pandas as pd
import os
import random
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
import hashlib
import time
import fcntl


DATA_DIR = 'data'
USER_DATA_FILE = os.path.join(DATA_DIR, 'zjudate.xlsx')
SURVEY_DATA_FILE = os.path.join(DATA_DIR, 'studentdata.xlsx')
os.makedirs(DATA_DIR, exist_ok=True)

verification_codes = {}

EMAIL_CONFIG = {
    'smtp_server': 'smtpdm.aliyun.com',
    'smtp_port': 80,
    'sender_email': 'info@crushzju.cn',
    'sender_password': '2442171329ABcd'
}     

def init_excel():
    if not os.path.exists(USER_DATA_FILE):
        df = pd.DataFrame(columns=['学号', '昵称', '密码', 'zju邮箱地址', '手机号', 
                                   '性别', '取向', '细分取向', '年级', '生日', '注册时间', '是否完成问卷'])
        df.to_excel(USER_DATA_FILE, index=False, engine="openpyxl")
    if not os.path.exists(SURVEY_DATA_FILE):
        # 82道题的列名
        columns = ['学号', '邮箱', '性别', '取向', '细分取向'] + [f'Q{i+1}' for i in range(82)]
        df = pd.DataFrame(columns=columns)
        df.to_excel(SURVEY_DATA_FILE, index=False, engine="openpyxl")

init_excel()

def hash_password(pwd):
    return hashlib.sha256(pwd.encode()).hexdigest()

def get_user_count():
    try:
        return len(pd.read_excel(USER_DATA_FILE, engine="openpyxl"))
    except:
        return 0

def check_completed(student_id):
    try:
        df = pd.read_excel(SURVEY_DATA_FILE, engine="openpyxl")
        return str(student_id) in df['学号'].astype(str).values
    except:
        return False

def send_email(to_email, code):
    try:
        msg = MIMEText(f'Your verification code is: {code}', 'plain', 'utf-8')
        msg['From'] = 'info@crushzju.cn'
        msg['To'] = to_email
        msg['Subject'] = 'CrushZJU Verification Code'
        
        server = smtplib.SMTP('smtpdm.aliyun.com', 80)
        server.login('info@crushzju.cn', '2442171329ABcd')
        server.sendmail('info@crushzju.cn', to_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"failed: {e}")
        return False

def save_excel_safe(df, filename):
    """安全保存Excel，防止并发写入损坏"""
    with open(filename, 'a') as f:
        fcntl.flock(f, fcntl.LOCK_EX)  # 加锁
        try:
            df.to_excel(filename, index=False, engine='openpyxl')
        finally:
            fcntl.flock(f, fcntl.LOCK_UN)  # 解锁

app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = 'crushzju_secret_key_2024'
app.config['JSON_AS_ASCII'] = False


@app.route('/')
def index():
    return render_template('index.html', user_count=get_user_count())

@app.route('/register')
def register_page():
    if session.get('logged_in'):
        return redirect(url_for('index'))
    return render_template('register.html')

@app.route('/login')
def login_page():
    if session.get('logged_in'):
        return redirect(url_for('index'))
    return render_template('login.html')

@app.route('/welcome')
def welcome_page():
    if not session.get('logged_in'):
        return redirect(url_for('index'))
    return render_template('welcome.html')

@app.route('/ready')
def ready_page():
    if not session.get('logged_in'):
        return redirect(url_for('index'))
    return render_template('ready.html')

@app.route('/survey')
def survey_page():
    if not session.get('logged_in'):
        return redirect(url_for('index'))
    if check_completed(session.get('student_id')):
        return redirect(url_for('index'))
    return render_template('survey.html')

@app.route('/api/send_code', methods=['POST'])
def send_code():
    data = request.json
    student_id = data.get('student_id', '').strip()
    if not student_id:
        return jsonify({'success': False, 'message': '请输入学号'})
    
    email = f"{student_id}@zju.edu.cn"
    now = time.time()
    
    if email in verification_codes:
        if now - verification_codes[email].get('last_send', 0) < 60:
            remaining = 60 - (now - verification_codes[email]['last_send'])
            return jsonify({'success': False, 'message': f'请等待{int(remaining)}秒'})
    
    code = ''.join([str(random.randint(0,9)) for _ in range(6)])
    verification_codes[email] = {'code': code, 'timestamp': now, 'attempts': 0, 'last_send': now}
    
    if send_email(email, code):
        return jsonify({'success': True, 'message': '验证码已发送'})
    return jsonify({'success': False, 'message': '发送失败'})

@app.route('/api/verify_code', methods=['POST'])
def verify_code():
    data = request.json
    student_id = data.get('student_id', '').strip()
    input_code = data.get('code', '').strip()
    email = f"{student_id}@zju.edu.cn"
    now = time.time()
    
    if email not in verification_codes:
        return jsonify({'success': False, 'message': '请先获取验证码'})
    
    info = verification_codes[email]
    if now - info['timestamp'] > 900:
        del verification_codes[email]
        return jsonify({'success': False, 'message': '验证码已过期'})
    
    if info['attempts'] >= 10:
        del verification_codes[email]
        return jsonify({'success': False, 'message': '错误次数过多'})
    
    if input_code == info['code']:
        session['temp_student_id'] = student_id
        return jsonify({'success': True, 'message': '验证成功'})
    
    info['attempts'] += 1
    return jsonify({'success': False, 'message': f'验证码错误，还剩{10-info["attempts"]}次'})

@app.route('/api/register', methods=['POST'])
def register():
    data = request.json
    student_id = session.get('temp_student_id')
    if not student_id:
        return jsonify({'success': False, 'message': '请先验证邮箱'})
    
    nickname = data.get('nickname', '').strip()
    password = data.get('password', '')
    confirm = data.get('confirm_password', '')
    phone = data.get('phone', '').strip()
    gender = data.get('gender', '')
    orientation = data.get('orientation', '')
    sub = data.get('sub_orientation', '')
    grade = data.get('grade', '')
    birthday = data.get('birthday', '')
    
    if len(password) < 8:
        return jsonify({'success': False, 'message': '密码至少8位'})
    if password != confirm:
        return jsonify({'success': False, 'message': '密码不一致'})
    if not phone or len(phone) != 11:
        return jsonify({'success': False, 'message': '手机号格式错误'})
    
    df = pd.read_excel(USER_DATA_FILE, engine="openpyxl")
    if nickname in df['昵称'].values:
        return jsonify({'success': False, 'message': '昵称已被使用'})
    if student_id in df['学号'].astype(str).values:
        return jsonify({'success': False, 'message': '学号已注册'})
    
    # 读取用户数据，如果文件损坏则重新创建
    try:
        df = pd.read_excel(USER_DATA_FILE, engine="openpyxl")
    except Exception as e:
        print(f"读取文件失败，重新创建: {e}")
        df = pd.DataFrame(columns=['学号', '昵称', '密码', 'zju邮箱地址', '手机号', 
                                   '性别', '取向', '细分取向', '年级', '生日', '注册时间', '是否完成问卷'])
    
    # 检查昵称是否已存在
    if nickname in df['昵称'].values:
        return jsonify({'success': False, 'message': '昵称已被使用'})
    
    # 检查学号是否已注册
    if student_id in df['学号'].astype(str).values:
        return jsonify({'success': False, 'message': '学号已注册'})
    
    # 创建新用户
    new = pd.DataFrame([{
        '学号': student_id,
        '昵称': nickname,
        '密码': hash_password(password),
        'zju邮箱地址': f"{student_id}@zju.edu.cn",
        '手机号': phone,
        '性别': gender,
        '取向': orientation,
        '细分取向': sub if sub else '',
        '年级': grade,
        '生日': birthday,
        '注册时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '是否完成问卷': '否'
    }])
    
    # 合并并保存
    try:
        df = pd.concat([df, new], ignore_index=True)
        df.to_excel(USER_DATA_FILE, index=False, engine='openpyxl')
        print(f"Excel 保存成功")
    except Exception as e:
        print(f"Excel 保存失败: {e}")
        return jsonify({'success': False, 'message': f'保存失败: {str(e)}'})
    
    session['logged_in'] = True
    session['student_id'] = student_id
    session['nickname'] = nickname
    session.pop('temp_student_id', None)
    
    print(f"注册成功: {student_id}")
    return jsonify({'success': True, 'message': '注册成功'})

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    student_id = data.get('student_id', '').strip()
    password = data.get('password', '')
    
    if not student_id or not password:
        return jsonify({'success': False, 'message': '请输入学号和密码'})
    
    try:
        df = pd.read_excel(USER_DATA_FILE, engine="openpyxl")
        user = df[df['学号'].astype(str) == student_id]
        
        if len(user) == 0:
            return jsonify({'success': False, 'message': '学号不存在'})
        
        stored_password = user.iloc[0]['密码']
        if hash_password(password) != stored_password:
            return jsonify({'success': False, 'message': '密码错误'})
        
        session['logged_in'] = True
        session['student_id'] = student_id
        session['nickname'] = user.iloc[0]['昵称']
        
        return jsonify({'success': True, 'message': '登录成功'})
    except Exception as e:
        print(f"登录失败: {e}")
        return jsonify({'success': False, 'message': '登录失败'})

@app.route('/api/check_login', methods=['GET'])
def check_login():
    if session.get('logged_in'):
        return jsonify({'logged_in': True, 'completed': check_completed(session.get('student_id')), 'nickname': session.get('nickname')})
    return jsonify({'logged_in': False})

@app.route('/api/submit_survey', methods=['POST'])
def submit_survey():
    if not session.get('logged_in'):
        return jsonify({'success': False, 'message': '请先登录'})
    
    student_id = session.get('student_id')
    if check_completed(student_id):
        return jsonify({'success': False, 'message': '已完成问卷'})
    
    answers = request.json.get('answers', [])
    
    # 获取用户基本信息
    try:
        user_df = pd.read_excel(USER_DATA_FILE, engine="openpyxl")
        user_info = user_df[user_df['学号'].astype(str) == student_id].iloc[0]
        email = user_info['zju邮箱地址']
        gender = user_info['性别']
        orientation = user_info['取向']
        sub_orientation = user_info['细分取向'] if pd.notna(user_info['细分取向']) else ''
    except Exception as e:
        print(f"获取用户信息失败: {e}")
        return jsonify({'success': False, 'message': '获取用户信息失败'})
    
    # 82道题 + 5个基础列
    columns = ['学号', '邮箱', '性别', '取向', '细分取向'] + [f'Q{i+1}' for i in range(82)]
    new_row = [student_id, email, gender, orientation, sub_orientation] + answers
    
    # 读取或创建问卷数据文件
    if os.path.exists(SURVEY_DATA_FILE):
        df = pd.read_excel(SURVEY_DATA_FILE, engine="openpyxl")
        # 如果列数不对，重新创建
        if len(df.columns) != len(columns):
            df = pd.DataFrame(columns=columns)
    else:
        df = pd.DataFrame(columns=columns)
    
    # 添加新数据
    new_df = pd.DataFrame([new_row], columns=columns)
    df = pd.concat([df, new_df], ignore_index=True)
    df.to_excel(SURVEY_DATA_FILE, index=False, engine="openpyxl")
    
    # 更新用户表的问卷状态
    user_df = pd.read_excel(USER_DATA_FILE, engine="openpyxl")
    user_df.loc[user_df['学号'].astype(str) == student_id, '是否完成问卷'] = '是'
    user_df.to_excel(USER_DATA_FILE, index=False, engine="openpyxl")
    
    return jsonify({'success': True, 'message': '提交成功'})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'success': True})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)