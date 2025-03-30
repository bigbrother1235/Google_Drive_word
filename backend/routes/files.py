import os
import logging
import traceback
from flask import request, jsonify, make_response, send_file, abort, send_from_directory
import google.oauth2.credentials
import googleapiclient.discovery
from googleapiclient.http import MediaIoBaseDownload
from io import BytesIO
from utils import create_credentials_from_token
from config import FRONTEND_DIR, CLIENT_CONFIG

# 设置日志记录器
logger = logging.getLogger(__name__)

def register_routes(app):
    @app.route('/')
    def serve_root():
        taskpane_path = os.path.join(FRONTEND_DIR, 'taskpane.html')
        logger.debug(f"Attempting to serve: {taskpane_path}")
        logger.debug(f"File exists: {os.path.exists(taskpane_path)}")
        
        if not os.path.exists(taskpane_path):
            logger.error(f"File not found: {taskpane_path}")
            # 列出FRONTEND_DIR的内容以帮助调试
            if os.path.exists(FRONTEND_DIR):
                logger.debug(f"Contents of {FRONTEND_DIR}: {os.listdir(FRONTEND_DIR)}")
            return jsonify({"error": "Frontend file not found. Check server logs."}), 404
            
        return send_file(taskpane_path)

    @app.route('/taskpane.html')
    def serve_taskpane():
        taskpane_path = os.path.join(FRONTEND_DIR, 'taskpane.html')
        logger.debug(f"Attempting to serve taskpane.html: {taskpane_path}")
        logger.debug(f"File exists: {os.path.exists(taskpane_path)}")
        
        if not os.path.exists(taskpane_path):
            logger.error(f"File not found: {taskpane_path}")
            # 列出FRONTEND_DIR的内容以帮助调试
            if os.path.exists(FRONTEND_DIR):
                logger.debug(f"Contents of {FRONTEND_DIR}: {os.listdir(FRONTEND_DIR)}")
            return jsonify({"error": "taskpane.html not found. Check server logs."}), 404
            
        return send_file(taskpane_path)

    @app.route('/taskpane.js')
    def serve_taskpane_js():
        taskpane_js_path = os.path.join(FRONTEND_DIR, 'taskpane.js')
        logger.debug(f"Attempting to serve taskpane.js: {taskpane_js_path}")
        logger.debug(f"File exists: {os.path.exists(taskpane_js_path)}")
        
        if not os.path.exists(taskpane_js_path):
            logger.error(f"File not found: {taskpane_js_path}")
            if os.path.exists(FRONTEND_DIR):
                logger.debug(f"Contents of {FRONTEND_DIR}: {os.listdir(FRONTEND_DIR)}")
            return jsonify({"error": "taskpane.js not found. Check server logs."}), 404
            
        return send_file(taskpane_js_path)

    # 通用静态文件服务路由
    @app.route('/<path:filename>')
    def serve_static(filename):
        try:
            file_path = os.path.join(FRONTEND_DIR, filename)
            if not os.path.exists(file_path):
                logger.error(f"File not found: {file_path}")
                return jsonify({"error": f"File {filename} not found"}), 404
                
            logger.debug(f"Serving static file: {file_path}")
            return send_from_directory(FRONTEND_DIR, filename)
        except Exception as e:
            logger.error(f"Error serving static file {filename}: {str(e)}")
            return jsonify({"error": str(e)}), 500

    @app.route('/api/list-files', methods=['POST'])
    @app.route('/api/list-folders', methods=['POST'])
    def list_folders():
        """列出指定文件夹中的子文件夹"""
        try:
            # 获取请求的JSON数据
            token_info = request.get_json()
            if not token_info:
                return jsonify({"error": "缺少令牌信息"}), 400

            # 创建凭证
            credentials = create_credentials_from_token(token_info)
            drive_service = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)
            
            # 获取文件夹ID，默认为根目录
            folder_id = token_info.get('folder_id', 'root')
            sort_by = token_info.get('sort_by', 'name')
            
            # 构建查询 - 仅文件夹
            query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
            
            # 排序字段
            order_by = "name" if sort_by == 'name' else 'modifiedTime desc'
            
            # 执行查询
            results = drive_service.files().list(
                q=query,
                fields="files(id, name, modifiedTime)",
                orderBy=order_by
            ).execute()
            
            logger.info(f"获取文件夹列表成功，找到 {len(results.get('files', []))} 个文件夹")
            
            return jsonify(results)
            
        except Exception as e:
            logger.error(f"处理list-folders请求时发生错误: {e}")
            logger.error(traceback.format_exc())
            return jsonify({"error": str(e)}), 500

    @app.route('/api/search-files', methods=['POST'])
    def search_files():
        """搜索文件和文件夹"""
        try:
            token_info = request.json
            credentials = create_credentials_from_token(token_info)
            
            search_term = request.json.get('search_term', '')
            if not search_term:
                return jsonify({"error": "搜索词不能为空"}), 400
            
            drive_service = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)
            
            # 构建查询
            query = f"name contains '{search_term}' and trashed=false"
            
            results = drive_service.files().list(
                q=query,
                fields="files(id, name, mimeType, webViewLink, modifiedTime, iconLink, parents)",
                pageSize=30
            ).execute()
            
            logger.info(f"搜索文件成功，找到 {len(results.get('files', []))} 个结果")
            
            return jsonify(results)
        except Exception as e:
            logger.error(f"搜索文件时出错: {str(e)}")
            return jsonify({"error": str(e)}), 500

    @app.route('/api/list-folder-contents', methods=['POST'])
    def list_folder_contents():
        """列出指定文件夹中的所有内容（包括文件和子文件夹）"""
        try:
            token_info = request.json
            credentials = create_credentials_from_token(token_info)
            
            folder_id = token_info.get('folder_id', 'root')
            
            drive_service = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)
            
            # 构建查询 - 所有未删除的文件和文件夹
            query = f"'{folder_id}' in parents and trashed=false"
            
            results = drive_service.files().list(
                q=query,
                fields="files(id, name, mimeType, modifiedTime, iconLink, size, webViewLink)",
                orderBy="name"
            ).execute()
            
            logger.info(f"获取文件夹内容成功，找到 {len(results.get('files', []))} 个项目")
            
            return jsonify(results)
        except Exception as e:
            logger.error(f"获取文件夹内容时出错: {str(e)}")
            return jsonify({"error": str(e)}), 500

    @app.route('/api/file-details', methods=['POST'])
    def get_file_details():
        """获取文件详情"""
        try:
            token_info = request.json
            file_id = token_info.get('file_id')
            
            if not file_id:
                return jsonify({"error": "缺少file_id参数"}), 400
            
            credentials = create_credentials_from_token(token_info)
            drive_service = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)
            
            # 获取文件详情
            file = drive_service.files().get(
                fileId=file_id,
                fields="id,name,mimeType,size,webViewLink,thumbnailLink,modifiedTime,createdTime,description,webContentLink"
            ).execute()
            
            # 如果是Google文档，尝试导出预览内容
            content = None
            if file.get('mimeType') == 'application/vnd.google-apps.document':
                try:
                    content = drive_service.files().export(
                        fileId=file_id,
                        mimeType='text/html'
                    ).execute()
                    if isinstance(content, bytes):
                        content = content.decode('utf-8')
                except Exception as ex:
                    logger.warning(f"导出Google文档内容时出错: {str(ex)}")
            
            return jsonify({
                "file": file,
                "content": content
            })
        except Exception as e:
            logger.error(f"获取文件详情时出错: {str(e)}")
            return jsonify({"error": str(e)}), 500

    @app.route('/api/direct-file-content/<file_id>', methods=['GET'])
    def get_direct_file_content(file_id):
        """直接获取文件内容（无需iframe）"""
        try:
            # 从查询参数获取令牌
            token = request.args.get('token')
            if not token:
                return jsonify({"error": "未提供访问令牌"}), 401
            
            # 创建凭证
            credentials = google.oauth2.credentials.Credentials(
                token=token,
                token_uri="https://oauth2.googleapis.com/token",
                client_id=CLIENT_CONFIG['installed']['client_id'],
                client_secret=CLIENT_CONFIG['installed']['client_secret']
            )
            
            # 创建Drive服务
            drive_service = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)
            
            # 请求文件元数据
            file_metadata = drive_service.files().get(fileId=file_id, fields="name,mimeType").execute()
            file_mime = file_metadata.get('mimeType', '')
            file_name = file_metadata.get('name', 'downloaded_file')
            
            # 下载文件内容
            request = drive_service.files().get_media(fileId=file_id)
            file_content = BytesIO()
            downloader = MediaIoBaseDownload(file_content, request)
            
            done = False
            while not done:
                status, done = downloader.next_chunk()
            
            file_content.seek(0)
            
            # 设置适当的Content-Type和Content-Disposition
            response = make_response(file_content.read())
            response.headers['Content-Type'] = file_mime
            response.headers['Content-Disposition'] = f'inline; filename="{file_name}"'
            response.headers['Access-Control-Allow-Origin'] = '*'
            
            return response
            
        except Exception as e:
            logger.error(f"获取文件内容时出错: {str(e)}")
            logger.error(traceback.format_exc())
            return jsonify({"error": str(e)}), 500