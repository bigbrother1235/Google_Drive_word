// fileOperations.js - 文件列表、浏览、搜索相关功能

import { safeLog, updateStatus, handleError } from './utils.js';
import { previewFile } from './filePreview.js';

// 根据mime类型获取图标
function getFileIcon(mimeType) {
    if (mimeType === 'application/vnd.google-apps.folder') {
        return mimeTypeIcons.folder;
    } else if (mimeType.includes('image/')) {
        return mimeTypeIcons.image;
    } else if (mimeType.includes('audio/')) {
        return mimeTypeIcons.audio;
    } else if (mimeType.includes('video/')) {
        return mimeTypeIcons.video;
    } else if (mimeType === 'application/pdf') {
        return mimeTypeIcons.pdf;
    } else if (mimeType === 'application/vnd.google-apps.document') {
        return mimeTypeIcons.document;
    } else if (mimeType === 'application/vnd.google-apps.spreadsheet') {
        return mimeTypeIcons.spreadsheet;
    } else if (mimeType === 'application/vnd.google-apps.presentation') {
        return mimeTypeIcons.presentation;
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        return mimeTypeIcons.word;
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        return mimeTypeIcons.excel;
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
        return mimeTypeIcons.powerpoint;
    } else if (mimeType.includes('zip') || mimeType.includes('compressed') || mimeType.includes('archive')) {
        return mimeTypeIcons.archive;
    } else {
        return mimeTypeIcons.unknown;
    }
}

// 获取文件类型的CSS类
function getFileClass(mimeType) {
    if (mimeType === 'application/vnd.google-apps.folder') {
        return 'folder';
    } else if (mimeType.includes('image/')) {
        return 'image';
    } else if (
        mimeType === 'application/vnd.google-apps.document' || 
        mimeType === 'application/vnd.google-apps.spreadsheet' || 
        mimeType === 'application/vnd.google-apps.presentation' ||
        mimeType === 'application/pdf' ||
        mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    ) {
        return 'document';
    } else {
        return 'other';
    }
}

// 加载文件夹内容
function loadFolderContents(folderId) {
    if (!tokenInfo) {
        handleError(new Error('未授权，请先完成Google授权'));
        return;
    }

    updateStatus('正在加载文件夹内容...');
    
    // 显示加载指示器
    const filesContainer = document.getElementById('files-container');
    filesContainer.innerHTML = '<div class="empty-message"><span class="loading"></span> 正在加载...</div>';
    
    // 构建请求体
    const requestBody = {
        token: tokenInfo.token,
        refresh_token: tokenInfo.refresh_token,
        token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
        client_id: tokenInfo.client_id,
        scopes: tokenInfo.scopes,
        folder_id: folderId
    };
    
    fetch(`${API_BASE_URL}/api/list-folder-contents`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(errorData => {
                throw new Error(`服务器错误: ${errorData.error || '未知错误'}`);
            });
        }
        return response.json();
    })
    .then(data => {
        safeLog('获取文件夹内容成功', data);
        renderFolderContents(data.files || []);
        updateStatus(`已加载 ${data.files ? data.files.length : 0} 个项目`);
    })
    .catch(error => {
        if (error.message.includes('401')) {
            handleError(error, true);
        } else {
            handleError(error);
        }
    });
}

// 渲染文件夹内容
function renderFolderContents(files) {
    const container = document.getElementById('files-container');
    container.innerHTML = '';

    if (files.length === 0) {
        container.innerHTML = '<div class="empty-message">此文件夹为空</div>';
        return;
    }

    // 首先显示文件夹，然后显示文件
    const folders = files.filter(file => file.mimeType === 'application/vnd.google-apps.folder');
    const otherFiles = files.filter(file => file.mimeType !== 'application/vnd.google-apps.folder');
    
    // 排序：按名称升序
    folders.sort((a, b) => a.name.localeCompare(b.name));
    otherFiles.sort((a, b) => a.name.localeCompare(b.name));
    
    // 合并排序后的文件夹和文件
    const sortedFiles = [...folders, ...otherFiles];
    
    sortedFiles.forEach(file => {
        const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
        const fileIcon = getFileIcon(file.mimeType);
        const fileClass = getFileClass(file.mimeType);
        
        const fileItem = document.createElement('div');
        fileItem.className = `file-item ${fileClass}`;
        fileItem.setAttribute('data-id', file.id);
        fileItem.setAttribute('data-name', file.name);
        fileItem.setAttribute('data-mime-type', file.mimeType);
        if (file.webViewLink) {
            fileItem.setAttribute('data-web-link', file.webViewLink);
        }
        
        fileItem.innerHTML = `
            <div class="file-icon">${fileIcon}</div>
            <div class="file-name" title="${file.name}">${file.name}</div>
        `;
        
        if (isFolder) {
            fileItem.onclick = function() {
                navigateToFolder(file.id, file.name);
            };
        } else {
            fileItem.onclick = function() {
                previewFile(file);
            };
        }
        
        container.appendChild(fileItem);
    });
}

// 设置搜索UI
function setupSearchUI() {
    const searchButton = document.getElementById('search-button');
    const searchInput = document.getElementById('search-input');
    
    if (searchButton && searchInput) {
        // 搜索按钮点击
        searchButton.onclick = function() {
            const searchTerm = searchInput.value.trim();
            if (searchTerm) {
                searchFiles(searchTerm);
            }
        };
        
        // 搜索框回车键
        searchInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                const searchTerm = searchInput.value.trim();
                if (searchTerm) {
                    searchFiles(searchTerm);
                }
            }
        });
    }
}

// 搜索文件功能
function searchFiles(searchTerm) {
    if (!tokenInfo) {
        handleError(new Error('未授权，请先完成Google授权'));
        return;
    }
    
    // 更新UI为搜索模式
    isSearchMode = true;
    currentSearchTerm = searchTerm;
    
    // 显示加载状态
    const filesContainer = document.getElementById('files-container');
    filesContainer.innerHTML = '<div class="empty-message"><span class="loading"></span> 正在搜索...</div>';
    
    // 构建请求体
    const requestBody = {
        token: tokenInfo.token,
        refresh_token: tokenInfo.refresh_token,
        token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
        client_id: tokenInfo.client_id,
        scopes: tokenInfo.scopes,
        search_term: searchTerm
    };
    
    // 发起API调用
    fetch(`${API_BASE_URL}/api/search-files`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(errorData => {
                throw new Error(`服务器错误: ${errorData.error || '未知错误'}`);
            });
        }
        return response.json();
    })
    .then(data => {
        safeLog('搜索文件成功', data);
        
        // 更新文件浏览器标题
        const breadcrumbContainer = document.getElementById('breadcrumb-container');
        breadcrumbContainer.innerHTML = `
            <span class="breadcrumb-item" data-id="root">根目录</span>
            <span class="breadcrumb-item search-result">搜索结果: "${searchTerm}"</span>
            <span class="breadcrumb-item close-search" title="返回浏览">×</span>
        `;
        
        // 添加返回浏览的点击事件
        document.querySelector('.close-search').onclick = function() {
            isSearchMode = false;
            // 返回当前文件夹视图
            const currentFolder = breadcrumbPath[breadcrumbPath.length - 1];
            loadFolderContents(currentFolder.id);
            renderBreadcrumb();
        };
        
        // 渲染搜索结果
        renderSearchResults(data.files || []);
        updateStatus(`找到 ${data.files ? data.files.length : 0} 个匹配项`);
    })
    .catch(error => {
        if (error.message.includes('401')) {
            handleError(error, true);
        } else {
            handleError(error);
        }
    });
}

// 渲染搜索结果
function renderSearchResults(files) {
    const container = document.getElementById('files-container');
    container.innerHTML = '';

    if (files.length === 0) {
        container.innerHTML = '<div class="empty-message">没有找到匹配的文件</div>';
        return;
    }

    // 首先显示文件夹，然后显示文件
    const folders = files.filter(file => file.mimeType === 'application/vnd.google-apps.folder');
    const otherFiles = files.filter(file => file.mimeType !== 'application/vnd.google-apps.folder');
    
    // 排序：按名称升序
    folders.sort((a, b) => a.name.localeCompare(b.name));
    otherFiles.sort((a, b) => a.name.localeCompare(b.name));
    
    // 合并排序后的文件夹和文件
    const sortedFiles = [...folders, ...otherFiles];
    
    sortedFiles.forEach(file => {
        const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
        const fileIcon = getFileIcon(file.mimeType);
        const fileClass = getFileClass(file.mimeType);
        
        const fileItem = document.createElement('div');
        fileItem.className = `file-item ${fileClass}`;
        fileItem.setAttribute('data-id', file.id);
        fileItem.setAttribute('data-name', file.name);
        fileItem.setAttribute('data-mime-type', file.mimeType);
        if (file.webViewLink) {
            fileItem.setAttribute('data-web-link', file.webViewLink);
        }
        
        fileItem.innerHTML = `
            <div class="file-icon">${fileIcon}</div>
            <div class="file-name" title="${file.name}">${file.name}</div>
        `;
        
        if (isFolder) {
            fileItem.onclick = function() {
                navigateToFolder(file.id, file.name);
                isSearchMode = false;
            };
        } else {
            fileItem.onclick = function() {
                previewFile(file);
            };
        }
        
        container.appendChild(fileItem);
    });
}

// 导航到指定文件夹
function navigateToFolder(folderId, folderName) {
    // 更新面包屑导航
    const currentIndex = breadcrumbPath.findIndex(item => item.id === folderId);
    
    if (currentIndex >= 0) {
        // 如果在路径中，则截断到该位置
        breadcrumbPath.splice(currentIndex + 1);
    } else {
        // 如果不在路径中，则添加到路径末尾
        breadcrumbPath.push({ id: folderId, name: folderName });
    }
    
    // 重新渲染面包屑
    renderBreadcrumb();
    
    // 加载文件夹内容
    loadFolderContents(folderId);
}

// 渲染面包屑导航
function renderBreadcrumb() {
    const container = document.getElementById('breadcrumb-container');
    container.innerHTML = '';
    
    breadcrumbPath.forEach((item, index) => {
        const breadcrumbItem = document.createElement('span');
        breadcrumbItem.className = 'breadcrumb-item';
        breadcrumbItem.textContent = item.name;
        breadcrumbItem.setAttribute('data-id', item.id);
        
        // 除最后一项外，其他项可点击
        if (index < breadcrumbPath.length - 1) {
            breadcrumbItem.onclick = function() {
                navigateToFolder(item.id, item.name);
            };
        }
        
        container.appendChild(breadcrumbItem);
    });
}

export { loadFolderContents, navigateToFolder, renderBreadcrumb, setupSearchUI, searchFiles, renderFolderContents, renderSearchResults, getFileIcon, getFileClass };