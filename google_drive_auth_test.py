import os
import json
import flask
import webbrowser
from flask import Flask, request, redirect, render_template_string
import google.oauth2.credentials
import google_auth_oauthlib.flow
import googleapiclient.discovery
import traceback

# 启用日志
import logging
logging.basicConfig(level=logging.DEBUG)
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
# 允许HTTP传输OAuth信息（仅用于本地测试）
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

app = Flask(__name__)
app.secret_key = os.urandom(24)

# 配置
CLIENT_SECRETS_FILE = 'client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

@app.route('/')
def index():
    """首页"""
    return render_template_string('''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Google Drive API调试工具</title>
            <style>
                body { font-family: Arial; max-width: 800px; margin: 0 auto; padding: 20px; }
                .btn { background: #4285F4; color: white; padding: 10px 20px; text-decoration: none; display: inline-block; border-radius: 4px; }
                pre { background: #f5f5f5; padding: 10px; border-radius: 4px; overflow-x: auto; }
                .card { border: 1px solid #ddd; border-radius: 4px; padding: 15px; margin: 15px 0; }
            </style>
        </head>
        <body>
            <h1>Google Drive API调试工具</h1>
            
            <div class="card">
                <h3>1. 环境信息</h3>
                <ul>
                    <li>Flask版本: {{ flask_version }}</li>
                    <li>client_secret.json: {{ client_secret_exists }}</li>
                </ul>
            </div>
            
            <div class="card">
                <h3>2. 开始授权测试</h3>
                <p>点击下面的按钮开始Google Drive授权流程:</p>
                <a href="/authorize" class="btn">开始授权</a>
            </div>
        </body>
        </html>
    ''', flask_version=flask.__version__,
         client_secret_exists='已找到' if os.path.exists(CLIENT_SECRETS_FILE) else '未找到')

@app.route('/authorize')
def authorize():
    """准备授权URL和状态"""
    try:
        # 记录client_secret.json内容
        if os.path.exists(CLIENT_SECRETS_FILE):
            with open(CLIENT_SECRETS_FILE, 'r') as f:
                client_data = json.load(f)
                if 'web' in client_data:
                    app.logger.info(f"客户端ID: {client_data['web'].get('client_id', 'N/A')[:5]}...")
                    app.logger.info(f"重定向URIs: {client_data['web'].get('redirect_uris', [])}")
                else:
                    app.logger.warning("客户端配置中未找到'web'部分")
        
        # 创建OAuth流程
        flow = google_auth_oauthlib.flow.Flow.from_client_secrets_file(
            CLIENT_SECRETS_FILE, scopes=SCOPES)
        flow.redirect_uri = 'http://localhost:8080/oauth2callback'
        
        # 生成授权URL
        authorization_url, state = flow.authorization_url(
            access_type='offline',
            include_granted_scopes='true',
            prompt='consent'
        )
        
        # 存储状态
        flask.session['state'] = state
        app.logger.info(f"生成授权URL，状态: {state}")
        
        # 重定向到授权页面
        return redirect(authorization_url)
    except Exception as e:
        tb = traceback.format_exc()
        app.logger.error(f"授权URL生成错误: {str(e)}\n{tb}")
        return render_template_string('''
            <h1>错误</h1>
            <p>生成授权URL时出现错误:</p>
            <pre>{{ error }}</pre>
            <pre>{{ traceback }}</pre>
            <p><a href="/">返回首页</a></p>
        ''', error=str(e), traceback=tb)

@app.route('/oauth2callback')
def oauth2callback():
    """处理OAuth回调"""
    try:
        # 记录回调URL
        callback_url = request.url
        masked_url = callback_url.replace(request.args.get('code', ''), '***') if 'code' in request.args else callback_url
        app.logger.info(f"收到回调: {masked_url}")
        
        # 检索状态
        state = flask.session.get('state')
        if not state:
            raise ValueError("无法从会话中获取state，可能是会话过期或cookie问题")
        
        # 比较状态值
        if state != request.args.get('state'):
            raise ValueError(f"状态不匹配: 期望 {state}, 获得 {request.args.get('state')}")
        
        # 创建OAuth流程
        flow = google_auth_oauthlib.flow.Flow.from_client_secrets_file(
            CLIENT_SECRETS_FILE, scopes=SCOPES, state=state)
        flow.redirect_uri = 'http://localhost:8080/oauth2callback'
        
        # 交换授权码获取令牌
        flow.fetch_token(authorization_response=request.url)
        
        # 获取凭证
        credentials = flow.credentials
        
        # 尝试使用令牌访问Drive API
        drive = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)
        files = drive.files().list(pageSize=5).execute()
        
        # 显示成功页面
        return render_template_string('''
            <html>
            <head>
                <title>授权成功</title>
                <style>
                    body { font-family: Arial; max-width: 800px; margin: 0 auto; padding: 20px; }
                    .success { color: green; }
                    pre { background: #f5f5f5; padding: 10px; border-radius: 4px; }
                </style>
            </head>
            <body>
                <h1 class="success">✓ 授权成功!</h1>
                <p>您已成功完成Google Drive API授权。</p>
                
                <h3>令牌信息</h3>
                <pre>
令牌: {{ token }}...
刷新令牌: {{ has_refresh_token }}
过期时间: {{ expiry }}
范围: {{ scopes }}
                </pre>
                
                <h3>Files API测试 (前5个文件)</h3>
                <ul>
                {% for file in files %}
                    <li>{{ file.name }} ({{ file.id }})</li>
                {% endfor %}
                </ul>
                
                <p><a href="/">返回首页</a></p>
            </body>
            </html>
        ''', token=credentials.token[:10], 
            has_refresh_token='已获取' if credentials.refresh_token else '未获取',
            expiry=credentials.expiry,
            scopes=credentials.scopes,
            files=files.get('files', []))
    
    except Exception as e:
        tb = traceback.format_exc()
        app.logger.error(f"OAuth回调处理错误: {str(e)}\n{tb}")
        return render_template_string('''
            <h1>授权过程中出错</h1>
            <p>错误信息:</p>
            <pre>{{ error }}</pre>
            <h3>详细堆栈跟踪:</h3>
            <pre>{{ traceback }}</pre>
            <h3>请求信息:</h3>
            <pre>URL: {{ url }}
参数: {{ args }}</pre>
            <p><a href="/">返回首页</a></p>
        ''', error=str(e), traceback=tb, url=request.url, args=dict(request.args))

if __name__ == '__main__':
    print("\n========== Google Drive API调试工具 ==========\n")
    
    # 检查client_secret.json
    if not os.path.exists(CLIENT_SECRETS_FILE):
        print(f"错误: 未找到{CLIENT_SECRETS_FILE}文件!")
        print(f"请从Google Cloud Console下载OAuth客户端凭证并保存为{CLIENT_SECRETS_FILE}\n")
        exit(1)
    
    # 自动打开浏览器
    webbrowser.open('http://localhost:8080')
    
    print("服务器已启动，请在浏览器中继续操作...\n")
    app.run('localhost', 8080, debug=True)