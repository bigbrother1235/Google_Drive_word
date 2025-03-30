// auth.js - 处理认证相关功能

import { safeLog, updateStatus, handleError } from './utils.js';

// 显示令牌刷新提示
function showTokenRefreshModal() {
    const modal = document.getElementById('token-refresh-modal');
    if (modal) {
        modal.style.display = 'block';
        
        // 添加刷新按钮事件
        const refreshButton = document.getElementById('refresh-token-button');
        if (refreshButton) {
            refreshButton.onclick = function() {
                modal.style.display = 'none';
                handleAuthClick();
            };
        }
    }
}

// 刷新令牌
function setupTokenRefresh() {
    // 清除任何现有的超时
    if (tokenRefreshTimeout) {
        clearTimeout(tokenRefreshTimeout);
    }
    
    // 设置新的超时 - 在令牌过期前显示提醒
    tokenRefreshTimeout = setTimeout(function() {
        showTokenRefreshModal();
    }, TOKEN_REFRESH_WARNING);
    
    safeLog('令牌刷新定时器已设置');
}

// 更新UI状态
function updateUIState(isAuthorized) {
    const authSection = document.getElementById('auth-section');
    const browserSection = document.getElementById('browser-section');
    const authButton = document.getElementById('auth-button');
    
    if (isAuthorized) {
        // 显示已授权状态
        authButton.textContent = '重新授权';
        browserSection.style.display = 'block';
        updateStatus('已成功授权Google Drive');
        
        // 设置令牌刷新提醒
        setupTokenRefresh();
        
        // 加载根目录内容
        loadFolderContents('root');
    } else {
        // 显示未授权状态
        authButton.textContent = '授权 Google Drive';
        browserSection.style.display = 'none';
        updateStatus('请点击上方按钮授权访问Google Drive');
    }
}

// 处理授权按钮点击
function handleAuthClick() {
    updateStatus('正在获取授权URL...');
    
    fetch(`${API_BASE_URL}/api/get-auth-url`, {
        method: 'GET',
        headers: {
            'Accept': 'application/json'
        }
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP错误! 状态: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        safeLog('获取授权URL成功', data);
        
        // 修改：不再使用localStorage存储state
        // 而是将整个授权URL显示给用户
        const authUrl = data.auth_url;
        
        // 更新状态显示授权链接
        const statusBox = document.getElementById('auth-status');
        statusBox.innerHTML = `
            <div>请点击下面的链接完成授权:</div>
            <div style="margin:10px 0;">
                <a href="${authUrl}" target="_blank" class="auth-link">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 48 48" style="vertical-align: middle; margin-right: 8px;">
                        <path fill="#FFC107" d="M43.611 20.083H42V20H24v8h11.303c-1.649 4.657-6.08 8-11.303 8c-6.627 0-12-5.373-12-12s5.373-12 12-12c3.059 0 5.842 1.154 7.961 3.039l5.657-5.657C34.046 6.053 29.268 4 24 4C12.955 4 4 12.955 4 24s8.955 20 20 20s20-8.955 20-20c0-1.341-.138-2.65-.389-3.917z"/>
                        <path fill="#FF3D00" d="m6.306 14.691l6.571 4.819C14.655 15.108 18.961 12 24 12c3.059 0 5.842 1.154 7.961 3.039l5.657-5.657C34.046 6.053 29.268 4 24 4C16.318 4 9.656 8.337 6.306 14.691z"/>
                        <path fill="#4CAF50" d="M24 44c5.166 0 9.86-1.977 13.409-5.192l-6.19-5.238A11.91 11.91 0 0 1 24 36c-5.202 0-9.619-3.317-11.283-7.946l-6.522 5.025C9.505 39.556 16.227 44 24 44z"/>
                        <path fill="#1976D2" d="M43.611 20.083H42V20H24v8h11.303a12.04 12.04 0 0 1-4.087 5.571l.003-.002l6.19 5.238C36.971 39.205 44 34 44 24c0-1.341-.138-2.65-.389-3.917z"/>
                    </svg>
                    授权Google Drive
                </a>
            </div>
            <div style="margin-top:16px; color:#666; background-color:#f8f9fa; padding:12px; border-radius:6px; border-left:4px solid #4285F4;">
                授权完成后，请复制浏览器地址栏中的完整URL，然后粘贴到下面的输入框中:
            </div>
            <div style="margin-top:16px;">
                <input type="text" id="callback-url" placeholder="粘贴授权后的URL..." style="width:100%; padding:10px; border-radius:6px; border:1px solid #dadce0;">
                <button id="process-callback" style="margin-top:12px; background-color:#4285F4; color:white; padding:10px 16px; border-radius:6px; font-weight:500;">处理授权回调</button>
            </div>
        `;
        
        // 添加回调处理按钮事件
        document.getElementById('process-callback').onclick = function() {
            const callbackUrl = document.getElementById('callback-url').value;
            if (!callbackUrl) {
                updateStatus('请粘贴授权后的URL', true);
                return;
            }
            
            try {
                // 从粘贴的URL中提取code和state参数
                const url = new URL(callbackUrl);
                const code = url.searchParams.get('code');
                const state = url.searchParams.get('state');
                
                if (!code || !state) {
                    updateStatus('无效的回调URL，未找到code或state参数', true);
                    return;
                }
                
                // 处理授权码
                processAuthCode(code, state);
            } catch (e) {
                updateStatus('无效的URL格式', true);
            }
        };
    })
    .catch(handleError);
}

// 处理授权码的函数
function processAuthCode(code, state) {
    updateStatus('正在处理授权...');
    
    // 发送授权码到后端交换令牌
    fetch(`${API_BASE_URL}/api/auth-callback`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ code, state })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP错误! 状态: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        safeLog('授权成功', data);
        
        // 保存令牌信息
        tokenInfo = data;
        localStorage.setItem('tokenInfo', JSON.stringify(tokenInfo));
        
        updateStatus('授权成功! 正在加载文件...');
        updateUIState(true);
    })
    .catch(handleError);
}

export { handleAuthClick, processAuthCode, updateUIState, setupTokenRefresh, showTokenRefreshModal };