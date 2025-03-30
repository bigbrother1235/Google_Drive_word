import logging
import traceback
from flask import request, jsonify, make_response
import google.oauth2.credentials
import google_auth_oauthlib.flow
import google.auth.transport.requests
from config import CLIENT_CONFIG, SCOPES, REDIRECT_URI
from utils import create_credentials_from_token, get_user_info

logger = logging.getLogger(__name__)

def register_routes(app):
    @app.route('/api/get-auth-url', methods=['GET', 'OPTIONS'])
    def get_auth_url():
        # 处理预检请求
        if request.method == 'OPTIONS':
            response = make_response()
            response.headers.add("Access-Control-Allow-Origin", "*")
            response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
            response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
            return response

        try:
            # 创建OAuth2流程
            flow = google_auth_oauthlib.flow.Flow.from_client_config(
                CLIENT_CONFIG, scopes=SCOPES
            )
            flow.redirect_uri = REDIRECT_URI
            
            # 生成授权URL
            authorization_url, state = flow.authorization_url(
                access_type='offline',
                include_granted_scopes='true',
                prompt='consent'  # 总是显示同意屏幕，确保获取refresh_token
            )
            
            logger.info(f"已生成授权URL，state={state}")
            
            # 创建响应并添加CORS头
            response = jsonify({
                'auth_url': authorization_url,
                'state': state
            })
            response.headers.add("Access-Control-Allow-Origin", "*")
            return response
        
        except Exception as e:
            logger.error(f"生成授权URL时出错: {str(e)}")
            error_response = jsonify({"error": str(e)})
            error_response.headers.add("Access-Control-Allow-Origin", "*")
            return error_response, 500

    @app.route('/api/auth-callback', methods=['POST'])
    def auth_callback():
        """处理OAuth2回调，交换授权码获取令牌"""
        try:
            data = request.json
            code = data.get('code')
            state = data.get('state')
            
            if not code or not state:
                return jsonify({"error": "缺少code或state参数"}), 400
            
            # 创建OAuth2流程
            flow = google_auth_oauthlib.flow.Flow.from_client_config(
                CLIENT_CONFIG, scopes=SCOPES
            )
            flow.redirect_uri = REDIRECT_URI
            
            # 交换授权码获取令牌
            flow.fetch_token(code=code)
            credentials = flow.credentials
            
            # 获取用户信息
            user_info = get_user_info(credentials)
            
            # 返回令牌和用户信息
            token_info = {
                'token': credentials.token,
                'refresh_token': credentials.refresh_token,
                'token_uri': credentials.token_uri,
                'client_id': credentials.client_id,
                'scopes': credentials.scopes,
                'expiry': credentials.expiry.isoformat() if credentials.expiry else None,
                'user_info': user_info
            }
            
            logger.info(f"授权成功，用户: {user_info.get('emailAddress', 'unknown')}")
            
            return jsonify(token_info)
        except Exception as e:
            logger.error(f"处理授权回调时出错: {str(e)}")
            logger.error(traceback.format_exc())
            return jsonify({"error": str(e)}), 500

    @app.route('/api/refresh-token', methods=['POST'])
    def refresh_token():
        """刷新访问令牌"""
        try:
            data = request.json
            refresh_token = data.get('refresh_token')
            client_id = data.get('client_id')
            
            if not refresh_token or not client_id:
                return jsonify({"error": "缺少刷新令牌或客户端ID"}), 400
            
            # 创建刷新令牌所需的凭据
            credentials = google.oauth2.credentials.Credentials(
                None,  # 没有访问令牌，因为我们要刷新它
                refresh_token=refresh_token,
                token_uri='https://oauth2.googleapis.com/token',
                client_id=client_id,
                client_secret=CLIENT_CONFIG['installed']['client_secret'] if 'installed' in CLIENT_CONFIG else CLIENT_CONFIG['web']['client_secret']
            )
            
            # 刷新访问令牌
            credentials.refresh(google.auth.transport.requests.Request())
            
            # 返回新的令牌信息
            token_info = {
                'access_token': credentials.token,
                'expiry': credentials.expiry.isoformat() if credentials.expiry else None
            }
            
            logger.info('成功刷新访问令牌')
            
            return jsonify(token_info)
        except Exception as e:
            logger.error(f"刷新令牌时出错: {str(e)}")
            logger.error(traceback.format_exc())
            return jsonify({"error": str(e)}), 500

    @app.route('/api/user-info', methods=['POST'])
    def get_user_info_endpoint():
        """获取用户信息"""
        try:
            token_info = request.json
            credentials = create_credentials_from_token(token_info)
            
            user_info = get_user_info(credentials)
            
            return jsonify(user_info)
        except Exception as e:
            logger.error(f"获取用户信息时出错: {str(e)}")
            return jsonify({"error": str(e)}), 500