// ==UserScript==
// @name               微信读书Word主题切换
// @name:zh-CN         微信读书Word主题切换
// @name:en            WeRead Word Theme Switcher
// @namespace          weread_word_style_themes
// @version            1.0.0
// @description        为微信读书Word风格添加主题切换功能
// @description:zh-CN  为微信读书Word风格添加主题切换功能
// @description:en     Add theme switching functionality to WeRead Word Style
// @author             AI Assistant
// @match              *://weread.qq.com/*
// @match              *://weread.qq.com/web/reader/*
// @grant              GM_addStyle
// @grant              GM_registerMenuCommand
// @grant              GM_setValue
// @grant              GM_getValue
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
    
    // 主题定义
    const themes = {
        // 经典白色主题
        'classic': {
            name: '经典白色',
            titleBarBg: '#f3f2f1',
            ribbonBg: '#f3f2f1',
            contentBg: '#ffffff',
            textColor: '#333333',
            accentColor: '#4472c4',
            borderColor: '#e1dfdd'
        },
        // 深色主题
        'dark': {
            name: '深色主题',
            titleBarBg: '#2d2d2d',
            ribbonBg: '#3c3c3c',
            contentBg: '#262626',
            textColor: '#e6e6e6',
            accentColor: '#4472c4',
            borderColor: '#505050'
        },
        // 彩色主题
        'colorful': {
            name: '彩色主题',
            titleBarBg: '#185abd',
            ribbonBg: '#2b579a',
            contentBg: '#ffffff',
            textColor: '#333333',
            accentColor: '#2b579a',
            borderColor: '#e1dfdd'
        },
        // 灰色主题
        'gray': {
            name: '灰色主题',
            titleBarBg: '#616161',
            ribbonBg: '#505050',
            contentBg: '#ffffff',
            textColor: '#333333',
            accentColor: '#505050',
            borderColor: '#e1dfdd'
        }
    };
    
    // 获取当前主题
    let currentTheme = GM_getValue('wordTheme', 'classic');
    
    // 应用主题
    function applyTheme(themeName) {
        if (!themes[themeName]) {
            themeName = 'classic';
        }
        
        const theme = themes[themeName];
        
        // 保存当前主题
        GM_setValue('wordTheme', themeName);
        currentTheme = themeName;
        
        // 创建主题CSS
        const themeCSS = `
            /* 标题栏 */
            #word-top-bar, .word-title-bar {
                background-color: ${theme.titleBarBg} !important;
                border-bottom-color: ${theme.borderColor} !important;
            }
            
            .word-title-bar .title {
                color: ${theme.textColor} !important;
            }
            
            /* 功能区 */
            .word-tabs, .word-ribbon {
                background-color: ${theme.ribbonBg} !important;
                border-color: ${theme.borderColor} !important;
            }
            
            .word-tabs .tab {
                color: ${theme.textColor} !important;
            }
            
            .word-tabs .tab.active {
                color: ${theme.accentColor} !important;
            }
            
            .word-tabs .tab.active::after {
                background-color: ${theme.accentColor} !important;
            }
            
            .ribbon-button .icon, .ribbon-button .label {
                color: ${theme.textColor} !important;
            }
            
            .ribbon-group {
                border-right-color: ${theme.borderColor} !important;
            }
            
            /* 内容区 */
            .readerContent, .app_content, .wr_whiteTheme {
                background-color: ${theme.contentBg} !important;
                border-color: ${theme.borderColor} !important;
            }
            
            .renderTargetContent, .bookContent, .readerChapter {
                color: ${theme.textColor} !important;
            }
            
            /* 状态栏 */
            #word-status-bar {
                background-color: ${theme.ribbonBg} !important;
                border-top-color: ${theme.borderColor} !important;
                color: ${theme.textColor} !important;
            }
            
            /* 页面背景 */
            body, html {
                background-color: ${theme.contentBg === '#ffffff' ? '#f5f5f5' : '#1e1e1e'} !important;
            }
        `;
        
        // 添加主题样式
        const styleId = 'word-theme-style';
        let styleElem = document.getElementById(styleId);
        
        if (styleElem) {
            styleElem.innerHTML = themeCSS;
        } else {
            styleElem = document.createElement('style');
            styleElem.id = styleId;
            styleElem.innerHTML = themeCSS;
            document.head.appendChild(styleElem);
        }
    }
    
    // 注册主题切换菜单
    function registerThemeMenus() {
        // 为每个主题注册菜单项
        Object.keys(themes).forEach(themeName => {
            const theme = themes[themeName];
            GM_registerMenuCommand(`Word主题: ${theme.name}`, () => {
                applyTheme(themeName);
            });
        });
    }
    
    // 监听DOM变化，确保主题应用到动态加载的元素
    function observeDOMChanges() {
        const observer = new MutationObserver(function(mutations) {
            mutations.forEach(function(mutation) {
                if (mutation.addedNodes.length) {
                    // 如果有新元素添加，确保主题应用
                    applyTheme(currentTheme);
                }
            });
        });
        
        // 开始观察文档变化
        observer.observe(document.documentElement, {
            childList: true,
            subtree: true
        });
    }
    
    // 页面加载完成后执行
    window.addEventListener('load', function() {
        applyTheme(currentTheme);
        registerThemeMenus();
        observeDOMChanges();
    });
    
    // 立即执行一次主题应用
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        setTimeout(function() {
            applyTheme(currentTheme);
            registerThemeMenus();
            observeDOMChanges();
        }, 500);
    }
})(); 