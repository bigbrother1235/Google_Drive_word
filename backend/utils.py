import logging
import re
import google.oauth2.credentials
import google.auth.transport.requests
import googleapiclient.discovery
from config import CLIENT_CONFIG, SCOPES

logger = logging.getLogger(__name__)

def create_credentials_from_token(token_info):
    """从令牌信息创建凭证对象"""
    try:
        return google.oauth2.credentials.Credentials(
            token=token_info.get('token'),
            refresh_token=token_info.get('refresh_token'),
            token_uri=token_info.get('token_uri', 'https://oauth2.googleapis.com/token'),
            client_id=token_info.get('client_id'),
            client_secret=CLIENT_CONFIG['installed']['client_secret'] if 'installed' in CLIENT_CONFIG else CLIENT_CONFIG['web']['client_secret'],
            scopes=token_info.get('scopes', SCOPES)
        )
    except Exception as e:
        logger.error(f"创建凭证时出错: {str(e)}")
        raise e

def get_user_info(credentials):
    """使用凭证获取用户信息"""
    try:
        drive_service = googleapiclient.discovery.build('drive', 'v3', credentials=credentials)
        
        about = drive_service.about().get(fields="user").execute()
        return about.get("user", {})
    except Exception as e:
        logger.error(f"获取用户信息时出错: {str(e)}")
        return {}

def sanitize_html(html):
    """简单的HTML清理函数"""
    # 移除所有脚本标签
    html = re.sub(r'<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>', '', html)
    
    # 移除可能的onclick等事件属性
    html = re.sub(r' on\w+="[^"]*"', '', html)
    
    return html