// ==UserScript==
// @name               微信读书Word 2021风格
// @name:zh-CN         微信读书Word 2021风格
// @name:en            WeRead Word 2021 Style
// @namespace          weread_word_style
// @version            1.0.0
// @description        将微信读书网页转换为Word 2021风格的界面
// @description:zh-CN  将微信读书网页转换为Word 2021风格的界面
// @description:en     Transform WeRead web page to Word 2021 style interface
// @author             AI Assistant
// @match              *://weread.qq.com/*
// @match              *://weread.qq.com/web/reader/*
// @grant              GM_addStyle
// @grant              GM_registerMenuCommand
// @run-at             document-start
// @license            MIT License
// @compatible         chrome 测试通过
// @compatible         edge 测试通过
// @compatible         firefox 测试通过
// @compatible         opera 未测试
// @compatible         safari 未测试
// ==/UserScript==

(function() {
    'use strict';
    
    // 添加Word 2021风格的CSS
    const wordStyleCSS = `
        /* 全局字体和颜色设置 */
        body, .readerContent, .app_content {
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif !important;
            background-color: #ffffff !important;
            color: #333333 !important;
        }
        
        /* 隐藏原微信读书的一些元素 */
        .readerTopBar, .navBar, .topBar {
            opacity: 0 !important;
            height: 0 !important;
            overflow: hidden !important;
        }
        
        /* 创建Word 2021风格的顶部区域 */
        #word-top-bar {
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            width: 100% !important;
            height: 110px !important;
            background-color: #f3f2f1 !important;
            z-index: 9999 !important;
            display: flex !important;
            flex-direction: column !important;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1) !important;
            border-bottom: 1px solid #e1dfdd !important;
        }
        
        /* Word 2021标题栏 */
        .word-title-bar {
            height: 30px !important;
            background-color: #f3f2f1 !important;
            display: flex !important;
            align-items: center !important;
            padding: 0 16px !important;
            border-bottom: 1px solid #e6e6e6 !important;
        }
        
        .word-title-bar .title {
            font-size: 12px !important;
            color: #666666 !important;
            flex-grow: 1 !important;
        }
        
        .word-title-bar .window-controls {
            display: flex !important;
        }
        
        .window-controls .control {
            width: 12px !important;
            height: 12px !important;
            margin-left: 8px !important;
            border-radius: 50% !important;
        }
        
        .window-controls .minimize {
            background-color: #ffbd2e !important;
        }
        
        .window-controls .maximize {
            background-color: #28c940 !important;
        }
        
        .window-controls .close {
            background-color: #ff5f57 !important;
        }
        
        /* Word 2021功能区选项卡 */
        .word-tabs {
            height: 30px !important;
            display: flex !important;
            align-items: center !important;
            padding: 0 16px !important;
            background-color: #f3f2f1 !important;
            border-bottom: 1px solid #e1dfdd !important;
        }
        
        .word-tabs .tab {
            padding: 6px 12px !important;
            font-size: 13px !important;
            color: #616161 !important;
            cursor: pointer !important;
            position: relative !important;
        }
        
        .word-tabs .tab.active {
            color: #4472c4 !important;
        }
        
        .word-tabs .tab.active::after {
            content: "" !important;
            position: absolute !important;
            bottom: -1px !important;
            left: 0 !important;
            width: 100% !important;
            height: 3px !important;
            background-color: #4472c4 !important;
        }
        
        /* Word 2021功能区 */
        .word-ribbon {
            height: 50px !important;
            background-color: #f3f2f1 !important;
            display: flex !important;
            align-items: center !important;
            padding: 0 16px !important;
            overflow-x: auto !important;
            overflow-y: hidden !important;
        }
        
        .ribbon-group {
            display: flex !important;
            height: 100% !important;
            padding: 0 8px !important;
            border-right: 1px solid #e1dfdd !important;
        }
        
        .ribbon-button {
            display: flex !important;
            flex-direction: column !important;
            align-items: center !important;
            justify-content: center !important;
            padding: 4px 8px !important;
            margin: 0 2px !important;
            height: 100% !important;
            background: transparent !important;
            border: none !important;
            cursor: pointer !important;
        }
        
        .ribbon-button:hover {
            background-color: rgba(0, 0, 0, 0.05) !important;
        }
        
        .ribbon-button .icon {
            font-size: 16px !important;
            color: #444 !important;
            margin-bottom: 2px !important;
        }
        
        .ribbon-button .label {
            font-size: 11px !important;
            color: #616161 !important;
        }
        
        /* 主内容区 */
        .readerContent, .app_content, .wr_whiteTheme {
            margin-top: 110px !important;
            padding: 40px !important;
            background-color: #ffffff !important;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1) !important;
            max-width: 800px !important;
            margin-left: auto !important;
            margin-right: auto !important;
            position: relative !important;
            border: 1px solid #e1dfdd !important;
        }
        
        /* 页面背景 */
        body, html {
            background-color: #f5f5f5 !important;
        }
        
        /* 阅读区域样式 */
        .renderTargetContent, .bookContent, .readerChapter {
            font-family: 'Calibri', 'Microsoft YaHei', sans-serif !important;
            font-size: 12pt !important;
            line-height: 1.5 !important;
            color: #333333 !important;
            text-align: justify !important;
        }
        
        /* 段落样式 */
        .renderTargetContent p, .bookContent p, .readerChapter p {
            margin-bottom: 12pt !important;
            text-indent: 2em !important;
        }
        
        /* 标题样式 */
        .renderTargetContent h1, .bookContent h1, .readerChapter h1 {
            font-size: 16pt !important;
            font-weight: bold !important;
            color: #2b579a !important;
            margin-top: 24pt !important;
            margin-bottom: 12pt !important;
        }
        
        .renderTargetContent h2, .bookContent h2, .readerChapter h2 {
            font-size: 14pt !important;
            font-weight: bold !important;
            color: #2b579a !important;
            margin-top: 18pt !important;
            margin-bottom: 9pt !important;
        }
        
        /* 页面边距指示 - 纸张效果 */
        .readerContent::before, .app_content::before, .wr_whiteTheme::before {
            content: "" !important;
            position: absolute !important;
            top: 0 !important;
            left: 0 !important;
            right: 0 !important;
            bottom: 0 !important;
            border: 1px solid #e1dfdd !important;
            background-color: #ffffff !important;
            z-index: -1 !important;
        }
        
        /* 页面阴影效果 */
        .readerContent::after, .app_content::after, .wr_whiteTheme::after {
            content: "" !important;
            position: absolute !important;
            top: 4px !important;
            left: 4px !important;
            right: -4px !important;
            bottom: -4px !important;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1) !important;
            z-index: -2 !important;
        }
        
        /* Word状态栏 */
        #word-status-bar {
            position: fixed !important;
            bottom: 0 !important;
            left: 0 !important;
            width: 100% !important;
            height: 22px !important;
            background-color: #f3f2f1 !important;
            border-top: 1px solid #e1dfdd !important;
            display: flex !important;
            align-items: center !important;
            padding: 0 16px !important;
            font-size: 12px !important;
            color: #616161 !important;
            z-index: 9999 !important;
        }
        
        .status-item {
            margin-right: 20px !important;
            display: flex !important;
            align-items: center !important;
        }
        
        .status-item .icon {
            margin-right: 4px !important;
            font-size: 12px !important;
        }
        
        /* 隐藏原有的一些元素 */
        .readerCatalog, .readerControls, .readerFooter .page, .readerFooter, .footerBar {
            display: none !important;
        }
        
        /* Word文档滚动条样式 */
        ::-webkit-scrollbar {
            width: 10px !important;
        }
        
        ::-webkit-scrollbar-track {
            background: #f1f1f1 !important;
        }
        
        ::-webkit-scrollbar-thumb {
            background: #c8c8c8 !important;
            border-radius: 5px !important;
        }
        
        ::-webkit-scrollbar-thumb:hover {
            background: #a8a8a8 !important;
        }
        
        /* 适配微信读书首页 */
        .wr_whiteTheme {
            margin-top: 110px !important;
        }
        
        /* 适配书架页 */
        .shelf_header, .navBar {
            opacity: 0 !important;
            height: 0 !important;
        }
        
        .shelf_container {
            margin-top: 110px !important;
        }
    `;
    
    // 添加样式到页面
    GM_addStyle(wordStyleCSS);
    
    // 创建Word 2021风格的UI元素
    function createWordUI() {
        // 创建Word顶部栏
        const wordTopBar = document.createElement('div');
        wordTopBar.id = 'word-top-bar';
        wordTopBar.innerHTML = `
            <div class="word-title-bar">
                <div class="title">微信读书 - Word</div>
                <div class="window-controls">
                    <div class="control minimize"></div>
                    <div class="control maximize"></div>
                    <div class="control close"></div>
                </div>
            </div>
            <div class="word-tabs">
                <div class="tab">文件</div>
                <div class="tab active">开始</div>
                <div class="tab">插入</div>
                <div class="tab">绘图</div>
                <div class="tab">设计</div>
                <div class="tab">布局</div>
                <div class="tab">引用</div>
                <div class="tab">审阅</div>
                <div class="tab">视图</div>
                <div class="tab">帮助</div>
            </div>
            <div class="word-ribbon">
                <div class="ribbon-group">
                    <button class="ribbon-button">
                        <span class="icon">📋</span>
                        <span class="label">剪切板</span>
                    </button>
                    <button class="ribbon-button">
                        <span class="icon">✂️</span>
                        <span class="label">剪切</span>
                    </button>
                    <button class="ribbon-button">
                        <span class="icon">📝</span>
                        <span class="label">复制</span>
                    </button>
                </div>
                <div class="ribbon-group">
                    <button class="ribbon-button">
                        <span class="icon">𝐁</span>
                        <span class="label">加粗</span>
                    </button>
                    <button class="ribbon-button">
                        <span class="icon">𝐼</span>
                        <span class="label">斜体</span>
                    </button>
                    <button class="ribbon-button">
                        <span class="icon">U̲</span>
                        <span class="label">下划线</span>
                    </button>
                </div>
                <div class="ribbon-group">
                    <button class="ribbon-button">
                        <span class="icon">⌘</span>
                        <span class="label">段落</span>
                    </button>
                    <button class="ribbon-button">
                        <span class="icon">⧉</span>
                        <span class="label">样式</span>
                    </button>
                </div>
                <div class="ribbon-group">
                    <button class="ribbon-button">
                        <span class="icon">🔍</span>
                        <span class="label">查找</span>
                    </button>
                    <button class="ribbon-button">
                        <span class="icon">🔄</span>
                        <span class="label">替换</span>
                    </button>
                </div>
                <div class="ribbon-group">
                    <button class="ribbon-button">
                        <span class="icon">⎙</span>
                        <span class="label">打印</span>
                    </button>
                </div>
            </div>
        `;
        
        // 创建Word状态栏
        const wordStatusBar = document.createElement('div');
        wordStatusBar.id = 'word-status-bar';
        wordStatusBar.innerHTML = `
            <div class="status-item">
                <span>页 1 / 共 1 页</span>
            </div>
            <div class="status-item">
                <span>字数: 100%</span>
            </div>
            <div class="status-item">
                <span>中文(中国)</span>
            </div>
        `;
        
        // 将元素添加到页面
        document.body.appendChild(wordTopBar);
        document.body.appendChild(wordStatusBar);
    }
    
    // 监听DOM变化，确保样式应用到动态加载的元素
    function observeDOMChanges() {
        const observer = new MutationObserver(function(mutations) {
            mutations.forEach(function(mutation) {
                if (mutation.addedNodes.length) {
                    // 如果有新元素添加，确保样式应用
                    applyWordStyles();
                }
            });
        });
        
        // 开始观察文档变化
        observer.observe(document.documentElement, {
            childList: true,
            subtree: true
        });
    }
    
    // 应用Word样式到特定元素
    function applyWordStyles() {
        // 更新状态栏的页码和进度信息
        updateStatusBar();
    }
    
    // 更新状态栏信息
    function updateStatusBar() {
        const statusBar = document.getElementById('word-status-bar');
        if (!statusBar) return;
        
        // 尝试获取当前页码和总页数
        let currentPage = 1;
        let totalPages = 1;
        
        // 尝试获取微信读书的页码信息
        const pageInfo = document.querySelector('.readerFooter .page');
        if (pageInfo) {
            const pageText = pageInfo.textContent;
            const match = pageText.match(/(\d+)\s*\/\s*(\d+)/);
            if (match) {
                currentPage = match[1];
                totalPages = match[2];
            }
        }
        
        // 更新状态栏
        const pageStatus = statusBar.querySelector('.status-item:first-child');
        if (pageStatus) {
            pageStatus.innerHTML = `<span>页 ${currentPage} / 共 ${totalPages} 页</span>`;
        }
    }
    
    // 注册菜单命令
    function registerMenuCommands() {
        GM_registerMenuCommand('切换Word风格', toggleWordStyle);
    }
    
    // 切换Word风格
    function toggleWordStyle() {
        const wordTopBar = document.getElementById('word-top-bar');
        const wordStatusBar = document.getElementById('word-status-bar');
        
        if (wordTopBar && wordStatusBar) {
            // 如果已存在，则移除
            wordTopBar.remove();
            wordStatusBar.remove();
            
            // 恢复原有的顶部栏
            const elements = document.querySelectorAll('.readerTopBar, .navBar, .topBar');
            elements.forEach(el => {
                el.style.opacity = '1';
                el.style.height = '';
                el.style.overflow = '';
            });
            
            // 恢复原有的内容区域
            const contentElements = document.querySelectorAll('.readerContent, .app_content, .wr_whiteTheme, .shelf_container');
            contentElements.forEach(el => {
                el.style.marginTop = '';
            });
        } else {
            // 如果不存在，则创建
            createWordUI();
        }
    }
    
    // 页面加载完成后执行
    window.addEventListener('load', function() {
        createWordUI();
        applyWordStyles();
        observeDOMChanges();
        registerMenuCommands();
    });
    
    // 立即执行一次样式应用
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        setTimeout(function() {
            createWordUI();
            applyWordStyles();
            observeDOMChanges();
            registerMenuCommands();
        }, 500);
    }
})(); 