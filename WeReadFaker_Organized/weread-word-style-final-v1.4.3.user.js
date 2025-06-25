// ==UserScript==
// @name         微信读书 - Word 2021风格
// @namespace    http://tampermonkey.net/
// @version      1.4.3
// @description  将微信读书网页版转换为Microsoft Word 2021风格，更加符合真实Word界面
// @author       AI助手
// @match        https://weread.qq.com/*
// @icon         https://res.wx.qq.com/a/wx_fed/weread/res/static/res/images/icon/icon_weread_logo_round.2c1c8.png
// @grant        none
// @license      MIT
// @run-at       document-start
// @compatible   chrome
// @compatible   firefox
// @compatible   edge
// ==/UserScript==

(function() {
    'use strict';
    
    let isStyleApplied = false;

    // SVG 图标定义
    const SVG_ICONS = {
        WORD_LOGO: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M21.5,17.5h-19C1.78,17.5,1,16.72,1,15.75V3.75C1,2.78,1.78,2,2.5,2h19C22.22,2,23,2.78,23,3.75v12 C23,16.72,22.22,17.5,21.5,17.5z M5,13.25h7V6.75H5V13.25z M12,13.25h7v-1.5h-7V13.25z M12,9.75h7V8.25h-7V9.75z M12,6.75h7V5.25h-7V6.75z"/></svg>',
        SEARCH: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="white" d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>',
        MINIMIZE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="white" d="M14 8v1H3V8h11z"/></svg>',
        MAXIMIZE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="white" d="M3 3v10h10V3H3zm9 9H4V4h8v8z"/></svg>',
        CLOSE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="white" d="M9.06 8l3.47-3.47a.75.75 0 00-1.06-1.06L8 6.94 4.53 3.47a.75.75 0 00-1.06 1.06L6.94 8 3.47 11.47a.75.75 0 001.06 1.06L8 9.06l3.47 3.47a.75.75 0 001.06-1.06L9.06 8z"/></svg>',
        SAVE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill="white" d="M4 2c-1.103 0-2 .897-2 2v12c0 1.103.897 2 2 2h12c1.103 0 2-.897 2-2V4c0-1.103-.897-2-2-2H4zm0 14V4h12l.002 12H4z"/><path fill="white" d="M9 5h2v4h4v2H9v-6z"/></svg>',
        UNDO: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill="white" d="M10 2c-4.411 0-8 3.589-8 8s3.589 8 8 8c3.212 0 6.007-1.897 7.391-4.636l-1.884-.79c-1.037 2.019-3.216 3.426-5.507 3.426-3.313 0-6-2.687-6-6s2.687-6 6-6c1.865 0 3.518.847 4.671 2.181l-3.671-.181V10h7V3h-1.08c-1.527-1.129-3.483-2-5.498-2z"/></svg>',
        REDO: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill="white" d="M10 2c4.411 0 8 3.589 8 8s-3.589 8-8 8c-3.212 0-6.007-1.897-7.391-4.636l1.884-.79c1.037 2.019 3.216 3.426 5.507 3.426 3.313 0 6-2.687 6-6s-2.687-6-6-6c-1.865 0-3.518.847-4.671 2.181l3.671-.181V10H3V3h1.08c1.527-1.129 3.483-2 5.498-2z"/></svg>',
        TOUCH_MOUSE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path fill="white" d="M10 2c-4.411 0-8 3.589-8 8s3.589 8 8 8c3.212 0 6.007-1.897 7.391-4.636l-1.884-.79c-1.037 2.019-3.216 3.426-5.507 3.426-3.313 0-6-2.687-6-6s2.687-6 6-6c1.865 0 3.518.847 4.671 2.181l-3.671-.181V10h7V3h-1.08c-1.527-1.129-3.483-2-5.498-2z"/></svg>', // Placeholder
        PASTE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M19 2h-4.5c-.32 0-.6.1-.85.3L12.5 3.5h-1c-.26-.2-.53-.3-.85-.3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-6 16H8v-2h5v2zm3-4H8v-2h8v2zm0-4H8V8h8v2z"/></svg>',
        CUT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M9 4c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zM4 16c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm5-12c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zm5 12c-1.1 0-2 .9-2 2s.9 2 2 2 2-.9 2-2-.9-2-2-2zM9 2v3h6V2H9zm-5 13v3h6v-3H4zm10 0v3h6v-3h-6zm0-11.23L15.23 3H19v4.23L17.77 8H14v-.23z"/></svg>',
        COPY: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>',
        FORMAT_PAINTER: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M18 4H6c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 14H6V6h12v12z"/><path fill="currentColor" d="M13 14H9v-2h4v2zM15 10H9V8h6v2z"/></svg>',
        BOLD: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M15.6 10.79c.97-1.1 1.6-2.58 1.6-4.29 0-3.53-2.73-5.21-6.1-5.21H3v18h10.42c3.47 0 6.09-1.74 6.09-5.32 0-2.45-1.16-4.04-2.91-4.38zM8.5 4.5h3c1.24 0 2.25.96 2.25 2.2s-1.01 2.2-2.25 2.2h-3V4.5zm4.5 13H8.5V14.5h4.5c1.24 0 2.25.96 2.25 2.2s-1.01 2.2-2.25 2.2z"/></svg>',
        ITALIC: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M10 4.5h4v2h-2.5l-3.5 12h2.5v2h-4v-2h2.5L12 6.5H9.5v-2z"/></svg>',
        UNDERLINE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M12 17c3.31 0 6-2.69 6-6V3h-2v8c0 2.21-1.79 4-4 4s-4-1.79-4-4V3H6v8c0 3.31 2.69 6 6 6zM5 19h14v2H5v-2z"/></svg>',
        STRIKETHROUGH: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M13.5 10.5H19v2h-5.5zm-8 0H11v2H5.5zM3 14h18v2H3zm-2-7h22v2H1z"/></svg>',
        TEXT_COLOR: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M0 20h24v4H0z"/><path fill="currentColor" d="M11 3L5.5 17h2.25l1.12-3h6.25l1.12 3h2.25L13 3zm-1.38 9L12 6.16 14.38 12H9.62z"/></svg>',
        HIGHLIGHT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M6 14l3 3l6-6l3 3V8l-6-6-6 6zm10.5-9.5l3 3L12 18H9l3-3 4.5-4.5z"/></svg>',
        BULLET_LIST: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M4 10.5c-.83 0-1.5.67-1.5 1.5s.67 1.5 1.5 1.5 1.5-.67 1.5-1.5-.67-1.5-1.5-1.5zM4 4.5c-.83 0-1.5.67-1.5 1.5S3.17 7.5 4 7.5s1.5-.67 1.5-1.5S4.83 4.5 4 4.5zm0 6c-.83 0-1.5.67-1.5 1.5s.67 1.5 1.5 1.5 1.5-.67 1.5-1.5-.67-1.5-1.5-1.5zM6 12h14v-2H6v2zm0-8v2h14V4H6zm0 8h14v-2H6v2z"/></svg>',
        NUMBER_LIST: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M2 17h2v.5H3v1h1v.5H2v1h3v-4H2v1zm1-9h1V4H2v1h1v3zm-1 9h1V4H2v1h1v3zm1 9h1V4H2v1h1v3zm1-9h1V4H2v1h1v3zm-1 9h1V4H2v1h1v3zm-1 9h1V4H2v1h1v3zM6 6h14v2H6V6zm0 12h14v2H6v-2zm0-6h14v2H6v-2z"/></svg>',
        ALIGN_LEFT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M15 15H3v2h12v-2zm0-8H3v2h12V7zM3 13h18v-2H3v2zm0 8h18v-2H3v2zM3 3v2h18V3H3z"/></svg>',
        ALIGN_CENTER: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M7 15v2h10v-2H7zm-4 6h18v-2H3v2zm0-8h18v-2H3v2zm4-6v2h10V7H7zM3 3v2h18V3H3z"/></svg>',
        ALIGN_RIGHT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M3 15h12v2H3zm0-8h12v2H3zm6 4h12V9H9zm0 8h12v-2H9zm-6-4h18v-2H3z"/></svg>',
        JUSTIFY: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M3 21h18v-2H3v2zm0-4h18v-2H3v2zm0-4h18v-2H3v2zm0-4h18V7H3v2zm0-4h18V3H3v2z"/></svg>',
        LINE_SPACING: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M16 11H8v2h8v-2zm-2 4H8v2h6v-2zm-4-8H8v2h2V7zm10 0h-8v2h8V7zm0 8h-8v2h8v-2z"/></svg>',
        HEADING1: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M17 11h-4V4H7v7H3v2h4v7h4v-7h4v7h4v-7h4v-2z"/></svg>',
        HEADING2: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M17 11h-4V4H7v7H3v2h4v7h4v-7h4v7h4v-7h4v-2z"/></svg>', // Placeholder
        NORMAL_TEXT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M14 17H4v-2h10v2zm0-4H4v-2h10v2zm0-4H4V7h10v2zm6 0v2h-2V9h2z"/></svg>',
        QUOTE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M6 17h3l2-4V7H5v6h3zm8 0h3l2-4V7h-6v6h3z"/></svg>',
        FIND: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>',
        REPLACE: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M16 10c0-1.1-.9-2-2-2h-3V5c0-1.1-.9-2-2-2s-2 .9-2 2v3H5c-1.1 0-2 .9-2 2s.9 2 2 2h3v3c0 1.1.9 2 2 2s2-.9 2-2v-3h3c1.1 0 2-.9 2-2zm-6 0v-2h2v2h-2zM9 13.5l1.5-1.5L9 10.5V13.5zm7.5 1.5L15 13.5 16.5 12v3z"/></svg>',
        SELECT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="currentColor" d="M3 13h18v-2H3v2zm0 4h18v-2H3v2zm0-8h18V7H3v2zm0-4h18V3H3v2z"/></svg>',
        
        // 新增翻页图标
        ARROW_LEFT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="white"><path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg>',
        ARROW_RIGHT: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="white"><path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/></svg>'
    };

    const WORD_STYLE_CSS = `
        /* Word 功能区容器 */
        #word-ribbon-container {
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            width: 100% !important;
            background-color: white !important;
            z-index: 10001 !important;
            font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif !important;
            box-shadow: 0 1px 4px rgba(0,0,0,0.1) !important;
            border-bottom: 1px solid #d0d0d0 !important;
        }

        /* 顶部蓝色条 */
        #word-top-bar {
            background-color: #2b579a !important;
            height: 32px !important;
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            padding: 0 8px !important; /* 调整内边距 */
            color: white !important;
            font-size: 12px !important;
            box-sizing: border-box !important;
        }

        /* 顶部左侧区域 */
        .word-top-left {
            display: flex !important;
            align-items: center !important;
            margin-left: 8px !important;
        }
        
        /* 顶部功能按钮 */
        .word-top-btn {
            width: 24px !important;
            height: 24px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            border-radius: 3px !important;
            cursor: pointer !important;
            margin-right: 4px !important;
        }
        
        .word-top-btn:hover {
            background-color: rgba(255,255,255,0.15) !important;
        }

        .word-top-btn svg {
            width: 16px !important;
            height: 16px !important;
            fill: white !important;
        }

        /* 自动保存开关 */
        .word-autosave {
            display: flex !important;
            align-items: center !important;
            margin-right: 12px !important;
        }

        .word-autosave-text {
            margin-right: 4px !important;
        }

        .word-autosave-toggle {
            width: 30px !important;
            height: 14px !important;
            background-color: #4c8bf5 !important;
            border-radius: 7px !important;
            position: relative !important;
            cursor: pointer !important;
        }

        .word-autosave-toggle::after {
            content: '';
            position: absolute !important;
            width: 12px !important;
            height: 12px !important;
            background-color: white !important;
            border-radius: 50% !important;
            top: 1px !important;
            right: 1px !important;
        }

        /* Word Logo */
        .word-logo-img {
            width: 20px !important;
            height: 20px !important;
            margin-right: 8px !important;
        }

        /* 文档标题 */
        .word-doc-title {
            font-weight: 400 !important;
            margin-right: 20px !important;
        }

        /* 顶部搜索框 */
        .word-search-bar {
            flex: 1 !important;
            max-width: 400px !important;
            height: 24px !important;
            border-radius: 3px !important;
            background-color: rgba(255, 255, 255, 0.2) !important;
            display: flex !important;
            align-items: center !important;
            padding: 0 8px !important;
            margin: 0 auto !important; /* 居中 */
        }

        .word-search-bar input {
            background: none !important;
            border: none !important;
            outline: none !important;
            color: white !important;
            font-size: 12px !important;
            flex: 1 !important;
            padding: 0 !important;
            margin-left: 8px !important;
        }
        
        .word-search-bar input::placeholder {
            color: rgba(255, 255, 255, 0.7) !important;
        }

        .word-search-icon-svg {
            width: 16px !important;
            height: 16px !important;
            fill: white !important;
            opacity: 0.8 !important;
        }

        /* 顶部右侧区域 */
        .word-top-right {
            display: flex !important;
            align-items: center !important;
            margin-left: auto !important;
        }

        /* 用户信息 */
        .word-user-info {
            display: flex !important;
            align-items: center !important;
            margin-right: 16px !important;
            cursor: pointer !important;
        }

        .word-user-avatar {
            width: 24px !important;
            height: 24px !important;
            border-radius: 50% !important;
            background-color: #8e918f !important;
            color: white !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-size: 14px !important;
            margin-right: 8px !important;
        }

        /* 窗口控制按钮 */
        .word-window-controls {
            display: flex !important;
            align-items: center !important;
            height: 100% !important;
        }

        .word-window-btn {
            width: 45px !important; /* 更宽的点击区域 */
            height: 100% !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            cursor: pointer !important;
        }
        
        .word-window-btn:hover {
            background-color: rgba(255,255,255,0.15) !important;
        }
        
        .word-window-btn.close-btn:hover {
            background-color: #e81123 !important;
        }
        
        .word-window-btn.close-btn:hover svg {
            fill: white !important;
        }

        .word-window-btn svg {
            width: 16px !important;
            height: 16px !important;
            fill: white !important;
        }

        /* 选项卡区域 */
        #word-ribbon-tabs {
            display: flex !important;
            padding: 0 8px !important;
            height: 30px !important; /* 调整高度 */
            background-color: #f3f2f1 !important;
            border-bottom: 1px solid #d0d0d0 !important;
        }
        
        .word-ribbon-tab {
            padding: 0 12px !important;
            height: 100% !important;
            display: flex !important;
            align-items: center !important;
            font-size: 13px !important; /* 调整字号 */
            color: #444 !important;
            cursor: pointer !important;
            border: 1px solid transparent !important;
            border-bottom: none !important;
            margin-right: 2px !important;
            transition: background-color 0.2s, border-color 0.2s !important;
        }
        
        .word-ribbon-tab:hover {
            background-color: #e5e5e5 !important;
        }
        
        .word-ribbon-tab.active {
            background-color: white !important;
            border-color: #d0d0d0 !important;
            border-bottom-color: white !important;
            color: #2b579a !important;
            font-weight: 600 !important;
        }

        /* 功能区内容 */
        #word-ribbon-content {
            height: 80px !important; /* 调整高度 */
            padding: 5px 10px !important;
            display: flex !important;
            background-color: white !important;
            align-items: flex-start !important; /* 顶部对齐 */
            overflow-x: auto !important;
            scrollbar-width: none !important; /* Firefox */
            -ms-overflow-style: none !important; /* IE and Edge */
            border-bottom: 1px solid #d0d0d0 !important;
        }
        
        #word-ribbon-content::-webkit-scrollbar {
            display: none !important; /* Chrome, Safari, Opera */
        }

        /* 功能区组 */
        .word-ribbon-group {
            display: flex !important;
            flex-direction: column !important;
            border-right: 1px solid #d0d0d0 !important;
            padding: 0px 5px !important;
            min-width: 60px !important;
            height: 100% !important;
            box-sizing: border-box !important;
            justify-content: flex-end !important; /* 底部对齐 */
        }
        
        .word-ribbon-group:last-child {
            border-right: none !important;
        }
        
        .word-ribbon-group-title {
            font-size: 10px !important;
            color: #666 !important;
            text-align: center !important;
            margin-top: auto !important; /* 底部对齐 */
            padding-bottom: 2px !important;
        }
        
        .word-ribbon-group-content {
            display: flex !important;
            flex-wrap: wrap !important;
            justify-content: center !important;
            align-items: center !important;
            flex-grow: 1 !important;
        }
        
        /* 功能区项 */
        .word-ribbon-item {
            height: 70px !important; /* 调整高度 */
            min-width: 40px !important;
            text-align: center !important;
            margin: 2px 1px !important;
            padding: 3px !important;
            border-radius: 3px !important;
            display: flex !important;
            flex-direction: column !important;
            align-items: center !important;
            justify-content: center !important;
            cursor: pointer !important;
            font-size: 11px !important;
            color: #444 !important;
            box-sizing: border-box !important;
        }
        
        .word-ribbon-item.large {
            height: 70px !important;
            width: 60px !important; /* 较大的按钮 */
            justify-content: center !important;
            padding-top: 10px !important;
        }

        .word-ribbon-item:hover {
            background-color: #e5e5e5 !important;
        }
        
        .word-ribbon-item-icon {
            font-size: 16px !important;
            height: 24px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            margin-bottom: 4px !important;
        }
        
        .word-ribbon-item-icon svg {
            width: 20px !important;
            height: 20px !important;
            fill: currentColor !important; /* 使用当前文字颜色 */
        }
        
        .word-ribbon-item-label {
            font-size: 10px !important;
            line-height: 1.2 !important;
        }

        /* 页面内容调整 */
        body {
            padding-top: 132px !important; /* 调整以适应新的功能区高度 (32 + 30 + 80 = 142px, 留一些余量) */
        }
        
        /* 隐藏原微信读书界面元素 */
        .readerTopBar, 
        .readerControls, 
        .readerFooter,
        .navBar,
        .footerBar,
        .app_content.reader {
            display: none !important;
        }
        
        /* 阅读区域修复 */
        #readerContent {
            margin-top: 0 !important;
            padding-top: 0 !important;
            transform: none !important;
        }
        
        /* 确保内容不被遮挡 */
        #root, .wr_whiteTheme, .app_content {
            margin-top: 0 !important;
        }

        /* 新增翻页按钮样式 */
        .word-page-turn-btn {
            position: fixed !important;
            top: 50% !important;
            transform: translateY(-50%) !important;
            width: 48px !important;
            height: 90px !important;
            background: linear-gradient(135deg, rgba(43,87,154,0.85) 60%, rgba(43,87,154,0.6) 100%) !important;
            color: white !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            cursor: pointer !important;
            z-index: 10002 !important;
            opacity: 0 !important;
            transition: opacity 0.25s, background 0.25s, box-shadow 0.25s !important;
            border-radius: 24px !important;
            box-shadow: 0 4px 16px 0 rgba(43,87,154,0.12), 0 1.5px 6px 0 rgba(0,0,0,0.08) !important;
            border: 1.5px solid rgba(255,255,255,0.18) !important;
            pointer-events: none;
        }

        .word-page-turn-btn.visible {
            opacity: 1 !important;
            pointer-events: auto;
        }

        .word-page-turn-btn svg {
            width: 36px !important;
            height: 36px !important;
            filter: drop-shadow(0 1px 2px rgba(0,0,0,0.12));
        }

        #word-page-turn-prev {
            left: 18px !important;
        }

        #word-page-turn-next {
            right: 18px !important;
        }

        .word-page-turn-btn:hover {
            background: linear-gradient(135deg, #2b579a 80%, #185abd 100%) !important;
            box-shadow: 0 8px 24px 0 rgba(43,87,154,0.22), 0 2px 8px 0 rgba(0,0,0,0.12) !important;
            border: 2px solid #fff !important;
        }
    `;
    
    function addStyles() {
        if (document.getElementById('word-style-css')) {
            return;
        }
        const styleElement = document.createElement('style');
        styleElement.id = 'word-style-css';
        styleElement.textContent = WORD_STYLE_CSS;
        if (document.head) {
            document.head.appendChild(styleElement);
        } else {
            const observer = new MutationObserver(function(mutations, obs) {
                if (document.head) {
                    document.head.appendChild(styleElement);
                    obs.disconnect();
                }
            });
            observer.observe(document.documentElement, { childList: true, subtree: true });
        }
    }
    
    function createWordInterface() {
        if (document.getElementById('word-ribbon-container')) {
            return;
        }
        
        const container = document.createElement('div');
        container.id = 'word-ribbon-container';
        
        const topBar = document.createElement('div');
        topBar.id = 'word-top-bar';
        
        const leftArea = document.createElement('div');
        leftArea.className = 'word-top-left';

        // 自动保存开关
        const autosave = document.createElement('div');
        autosave.className = 'word-autosave';
        autosave.innerHTML = `
            <div class="word-autosave-text">自动保存</div>
            <div class="word-autosave-toggle"></div>
        `;
        leftArea.appendChild(autosave);

        // 保存按钮
        const saveBtn = document.createElement('div');
        saveBtn.className = 'word-top-btn';
        saveBtn.innerHTML = SVG_ICONS.SAVE;
        leftArea.appendChild(saveBtn);

        // 撤销按钮
        const undoBtn = document.createElement('div');
        undoBtn.className = 'word-top-btn';
        undoBtn.innerHTML = SVG_ICONS.UNDO;
        leftArea.appendChild(undoBtn);

        // 重做按钮
        const redoBtn = document.createElement('div');
        redoBtn.className = 'word-top-btn';
        redoBtn.innerHTML = SVG_ICONS.REDO;
        leftArea.appendChild(redoBtn);

        // 触摸/鼠标模式切换
        const touchMouseBtn = document.createElement('div');
        touchMouseBtn.className = 'word-top-btn';
        touchMouseBtn.innerHTML = SVG_ICONS.TOUCH_MOUSE;
        leftArea.appendChild(touchMouseBtn);
        
        const logo = document.createElement('img');
        logo.className = 'word-logo-img';
        logo.src = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PHBhdGggZmlsbD0id2hpdGUiIGQ9Ik0yMS41LDE3LjVoLTE5QzEuNzgsMTcuNSwxLDE2LjcyLDEsMTUuNzVWMy43NUMxLDIuNzgsMS43OCwyLDIuNSwyaDE5QzIyLjIyLDIsMjMsMi43OCwyMywzLjc1djEyIEMyMywxNi43MiwyMi4yMiwxNy41LDIxLjUsMTcuNXogTTUsMTMuMjVoN1Y2Ljc1SDVWMTMuMjV6IE0xMiwxMy4yNWg3di0xLjVoLTdWMTMuMjV6IE0xMiw5Ljc1aDdWOC4yNWgtN1Y5Ljc1eiBNMTIsNi43NWg3VjUuMjVoLTdWNi43NXoiLz48L3N2Zz4=';
        
        const title = document.createElement('div');
        title.className = 'word-doc-title';
        title.textContent = '文档1 - Word';
        
        leftArea.appendChild(logo);
        leftArea.appendChild(title);
        
        // 搜索框
        const searchBar = document.createElement('div');
        searchBar.className = 'word-search-bar';
        searchBar.innerHTML = `
            ${SVG_ICONS.SEARCH}
            <input type="text" placeholder="搜索">
        `;
        topBar.appendChild(searchBar);

        const rightArea = document.createElement('div');
        rightArea.className = 'word-top-right';
        
        const userInfo = document.createElement('div');
        userInfo.className = 'word-user-info';
        userInfo.innerHTML = `
            <div class="word-user-avatar">O</div>
            <span>Odyssey Wang</span>
        `;
        rightArea.appendChild(userInfo);
        
        const windowControls = document.createElement('div');
        windowControls.className = 'word-window-controls';
        
        const minBtn = document.createElement('div');
        minBtn.className = 'word-window-btn';
        minBtn.innerHTML = SVG_ICONS.MINIMIZE;
        
        const maxBtn = document.createElement('div');
        maxBtn.className = 'word-window-btn';
        maxBtn.innerHTML = SVG_ICONS.MAXIMIZE;
        
        const closeBtn = document.createElement('div');
        closeBtn.className = 'word-window-btn close-btn'; /* 添加close-btn类 */
        closeBtn.innerHTML = SVG_ICONS.CLOSE;
        
        windowControls.appendChild(minBtn);
        windowControls.appendChild(maxBtn);
        windowControls.appendChild(closeBtn);
        
        rightArea.appendChild(windowControls);
        
        topBar.appendChild(leftArea);
        topBar.appendChild(rightArea);
        container.appendChild(topBar);
        
        const tabs = document.createElement('div');
        tabs.id = 'word-ribbon-tabs';
        
        const tabNames = ['文件', '开始', '插入', '绘图', '设计', '布局', '引用', '邮件', '审阅', '视图', '帮助'];
        
        tabNames.forEach((name, index) => {
            const tab = document.createElement('div');
            tab.className = 'word-ribbon-tab';
            if (index === 1) {
                tab.classList.add('active');
            }
            tab.textContent = name;
            tabs.appendChild(tab);
        });
        
        container.appendChild(tabs);
        
        const content = document.createElement('div');
        content.id = 'word-ribbon-content';
        
        // 剪贴板组
        const clipboardGroup = createRibbonGroup('剪贴板', [
            createRibbonItem('粘贴', SVG_ICONS.PASTE, 'large'), /* 较大的粘贴按钮 */
            createRibbonItem('剪切', SVG_ICONS.CUT),
            createRibbonItem('复制', SVG_ICONS.COPY),
            createRibbonItem('格式刷', SVG_ICONS.FORMAT_PAINTER)
        ]);
        
        // 字体组
        const fontGroup = createRibbonGroup('字体', [
            createRibbonItem('加粗', SVG_ICONS.BOLD),
            createRibbonItem('斜体', SVG_ICONS.ITALIC),
            createRibbonItem('下划线', SVG_ICONS.UNDERLINE),
            createRibbonItem('删除线', SVG_ICONS.STRIKETHROUGH),
            createRibbonItem('文本颜色', SVG_ICONS.TEXT_COLOR),
            createRibbonItem('高亮', SVG_ICONS.HIGHLIGHT)
        ]);

        // 段落组
        const paragraphGroup = createRibbonGroup('段落', [
            createRibbonItem('项目符号', SVG_ICONS.BULLET_LIST),
            createRibbonItem('编号', SVG_ICONS.NUMBER_LIST),
            createRibbonItem('左对齐', SVG_ICONS.ALIGN_LEFT),
            createRibbonItem('居中', SVG_ICONS.ALIGN_CENTER),
            createRibbonItem('右对齐', SVG_ICONS.ALIGN_RIGHT),
            createRibbonItem('两端对齐', SVG_ICONS.JUSTIFY),
            createRibbonItem('行距', SVG_ICONS.LINE_SPACING)
        ]);

        // 样式组
        const stylesGroup = createRibbonGroup('样式', [
            createRibbonItem('标题1', SVG_ICONS.HEADING1),
            createRibbonItem('标题2', SVG_ICONS.HEADING2),
            createRibbonItem('正文', SVG_ICONS.NORMAL_TEXT),
            createRibbonItem('引用', SVG_ICONS.QUOTE)
        ]);

        // 编辑组
        const editingGroup = createRibbonGroup('编辑', [
            createRibbonItem('查找', SVG_ICONS.FIND),
            createRibbonItem('替换', SVG_ICONS.REPLACE),
            createRibbonItem('选择', SVG_ICONS.SELECT)
        ]);

        content.appendChild(clipboardGroup);
        content.appendChild(fontGroup);
        content.appendChild(paragraphGroup);
        content.appendChild(stylesGroup);
        content.appendChild(editingGroup);

        container.appendChild(content);
        
        if (document.body) {
            document.body.insertBefore(container, document.body.firstChild);
            addTabsInteractivity();
        } else {
            const observer = new MutationObserver(function(mutations, obs) {
                if (document.body) {
                    document.body.insertBefore(container, document.body.firstChild);
                    addTabsInteractivity();
                    obs.disconnect();
                }
            });
            observer.observe(document.documentElement, { childList: true, subtree: true });
        }
    }
    
    function createRibbonGroup(title, items) {
        const group = document.createElement('div');
        group.className = 'word-ribbon-group';
        
        const content = document.createElement('div');
        content.className = 'word-ribbon-group-content';
        
        items.forEach(item => {
            content.appendChild(item);
        });
        
        const titleElement = document.createElement('div');
        titleElement.className = 'word-ribbon-group-title';
        titleElement.textContent = title;
        
        group.appendChild(content);
        group.appendChild(titleElement);
        
        return group;
    }
    
    function createRibbonItem(label, iconSVG, sizeClass = '') {
        const item = document.createElement('div');
        item.className = 'word-ribbon-item';
        if (sizeClass) {
            item.classList.add(sizeClass);
        }
        
        const iconElement = document.createElement('div');
        iconElement.className = 'word-ribbon-item-icon';
        iconElement.innerHTML = iconSVG; // 使用 SVG
        
        const labelElement = document.createElement('div');
        labelElement.className = 'word-ribbon-item-label';
        labelElement.textContent = label;
        
        item.appendChild(iconElement);
        item.appendChild(labelElement);
        
        return item;
    }
    
    function addTabsInteractivity() {
        const tabs = document.querySelectorAll('.word-ribbon-tab');
        if (!tabs || tabs.length === 0) return;
        
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                tabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
            });
        });
    }
    
    function hideOriginalElements() {
        const elementsToHide = [
            '.readerTopBar', 
            '.readerControls', 
            '.readerFooter',
            '.navBar',
            '.footerBar',
            '.app_content.reader' /* 确保隐藏整个微信读书阅读器部分，由我们的padding-top代替 */
        ];
        
        elementsToHide.forEach(selector => {
            const elements = document.querySelectorAll(selector);
            elements.forEach(el => {
                if (el) el.style.display = 'none';
            });
        });
    }
    
    function fixContentArea() {
        if (document.body) {
            document.body.style.paddingTop = '132px';
        }
        
        // 尝试找到微信读书的内容区域并进行调整
        const readerContent = document.querySelector('#readerContent'); // 修改选择器以适应实际情况
        if (readerContent) {
            readerContent.style.marginTop = '0';
            readerContent.style.paddingTop = '0';
            readerContent.style.transform = 'none';
        }

        // 进一步隐藏可能的其他遮挡元素
        const rootElement = document.querySelector('#root');
        if (rootElement) {
            rootElement.style.marginTop = '0';
        }
        const whiteTheme = document.querySelector('.wr_whiteTheme');
        if (whiteTheme) {
            whiteTheme.style.marginTop = '0';
        }
        const appContent = document.querySelector('.app_content');
        if (appContent) {
            appContent.style.marginTop = '0';
        }
    }
    
    function createPageTurnButtons() {
        if (document.getElementById('word-page-turn-prev')) {
            return; // 按钮已存在
        }
        const prevBtn = document.createElement('div');
        prevBtn.id = 'word-page-turn-prev';
        prevBtn.className = 'word-page-turn-btn';
        prevBtn.innerHTML = SVG_ICONS.ARROW_LEFT;
        prevBtn.addEventListener('click', () => {
            const originalPrevBtn = document.querySelector('.readerChapter_button_prev, .wr_readerPrevPage, .wr_pagePrev, .wr-btn-prev');
            if (originalPrevBtn) {
                originalPrevBtn.click();
            } else {
                const evt = new KeyboardEvent('keydown', { key: 'ArrowLeft', keyCode: 37, which: 37, bubbles: true });
                document.dispatchEvent(evt);
            }
        });
        const nextBtn = document.createElement('div');
        nextBtn.id = 'word-page-turn-next';
        nextBtn.className = 'word-page-turn-btn';
        nextBtn.innerHTML = SVG_ICONS.ARROW_RIGHT;
        nextBtn.addEventListener('click', () => {
            const originalNextBtn = document.querySelector('.readerChapter_button_next, .wr_readerNextPage, .wr_pageNext, .wr-btn-next');
            if (originalNextBtn) {
                originalNextBtn.click();
            } else {
                const evt = new KeyboardEvent('keydown', { key: 'ArrowRight', keyCode: 39, which: 39, bubbles: true });
                document.dispatchEvent(evt);
            }
        });
        const mainContent = document.querySelector('.app_content, .wr_whiteTheme, .readerContent, #readerContent, .readerChapterContent');
        if (mainContent) {
            mainContent.appendChild(prevBtn);
            mainContent.appendChild(nextBtn);
        } else {
            document.body.appendChild(prevBtn);
            document.body.appendChild(nextBtn);
        }
        // 鼠标靠近页面左右100px区域时显示按钮
        let isHovering = false;
        const showButtons = () => {
            prevBtn.classList.add('visible');
            nextBtn.classList.add('visible');
        };
        const hideButtons = () => {
            if (!isHovering) {
                prevBtn.classList.remove('visible');
                nextBtn.classList.remove('visible');
            }
        };
        [prevBtn, nextBtn].forEach(btn => {
            btn.addEventListener('mouseenter', () => {
                isHovering = true;
                showButtons();
            });
            btn.addEventListener('mouseleave', () => {
                isHovering = false;
                hideButtons();
            });
        });
        // 监听鼠标移动，靠近左右100px显示
        document.addEventListener('mousemove', (e) => {
            const w = window.innerWidth;
            if (e.clientX <= 100) {
                prevBtn.classList.add('visible');
            } else {
                if (!isHovering) prevBtn.classList.remove('visible');
            }
            if (e.clientX >= w - 100) {
                nextBtn.classList.add('visible');
            } else {
                if (!isHovering) nextBtn.classList.remove('visible');
            }
        });
    }
    
    function init() {
        if (isStyleApplied) return;
        
        try {
            addStyles();
            createWordInterface();
            // 始终尝试创建翻页按钮，具体的可见性由后续逻辑决定
            createPageTurnButtons();
            hideOriginalElements();
            fixContentArea();
            
            isStyleApplied = true;
            console.log('微信读书 - Word 2021风格 v1.4.3 初始化成功');
        } catch (error) {
            console.error('初始化Word风格时出错:', error);
        }
    }
    
    addStyles();
    
    document.addEventListener('DOMContentLoaded', init);
    window.addEventListener('load', init);
    
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        init();
    }
    
    const observer = new MutationObserver(function(mutations) {
        if (document.querySelector('#readerContent') && !isStyleApplied) { // 修改选择器
            init();
        }
    });
    
    observer.observe(document.documentElement, { childList: true, subtree: true });
    
    setTimeout(init, 1000);
    setTimeout(init, 2000);
})();