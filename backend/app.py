import os
import logging
import traceback
from flask import Flask, jsonify, request, send_from_directory, abort
from flask_cors import CORS

# 导入配置和工具模块
from config import FRONTEND_DIR, logger
from routes import auth, files, documents

# 创建Flask应用
app = Flask(__name__, static_folder=FRONTEND_DIR, static_url_path='/')
CORS(app, resources={
    r"/*": {
        "origins": "*",
        "allow_headers": ["Content-Type", "Authorization", "Access-Control-Allow-Credentials"],
        "supports_credentials": True
    }
})

# 日志记录所有请求
@app.before_request
def log_request_info():
    logger.debug('Request Headers: %s', request.headers)
    logger.debug('Request Path: %s', request.path)

# 注册各模块的路由
auth.register_routes(app)
files.register_routes(app)
documents.register_routes(app)

# 错误处理
@app.errorhandler(404)
def page_not_found(e):
    logger.error(f'404 error: {request.path}')
    return jsonify(error=str(e)), 404

@app.errorhandler(500)
def internal_server_error(e):
    logger.error(f'500 error: {str(e)}')
    return jsonify(error=str(e)), 500

if __name__ == '__main__':
    # 为本地开发使用自签名证书
    try:
        import ssl
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain('cert.pem', 'key.pem')
    except Exception as e:
        logger.error(f'证书加载错误: {e}')
        # 如果证书文件不存在，使用'adhoc'自动生成临时证书
        context = 'adhoc'
    
    logger.info("Flask服务器启动中，监听端口5001...")
    # 使用HTTPS运行Flask应用
    app.run(host='0.0.0.0', port=5001, ssl_context=context, debug=True)
    # Check if frontend directory exists before starting
if not os.path.exists(FRONTEND_DIR):
    logger.error(f"Frontend directory not found: {FRONTEND_DIR}")
    logger.error("Please make sure the frontend directory exists with required files")
    # List the project directory contents to help debugging
    parent_dir = os.path.dirname(CURRENT_DIR)
    logger.error(f"Contents of project directory {parent_dir}: {os.listdir(parent_dir)}")
    exit(1)
else:
    logger.info(f"Frontend directory exists: {FRONTEND_DIR}")
    logger.info(f"Frontend directory contents: {os.listdir(FRONTEND_DIR)}")