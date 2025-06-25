# 微信读书 - Word 2021风格 (v1.4.3)

这是一个将微信读书网页版转换为Microsoft Word 2021风格的Tampermonkey脚本，提供更专业、更舒适的阅读体验。

## 功能特点

- **Word 2021风格界面**：完整模拟Word 2021的功能区和标题栏
- **沉浸式阅读体验**：隐藏原有微信读书界面元素，专注内容
- **稳定可靠**：多重初始化机制确保在各种网络和浏览器环境下正常工作
- **轻量级**：优化代码结构，提高加载和执行速度
- **更符合Word视觉风格**：替换功能区图标为SVG，顶部蓝色条细节更完善

## 安装指南

1. 首先确保你已安装Tampermonkey浏览器扩展
   - [Chrome商店链接](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
   - [Firefox附加组件链接](https://addons.mozilla.org/zh-CN/firefox/addon/tampermonkey/)
   - [Edge插件链接](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)

2. 安装脚本
   - 方法1：直接点击[安装链接](weread-word-style-final-v1.4.3.user.js)
   - 方法2：复制`weread-word-style-final-v1.4.3.user.js`中的代码，在Tampermonkey中新建脚本并粘贴

3. 确保脚本已启用
   - 在Tampermonkey插件菜单中，确认"微信读书 - Word 2021风格"脚本处于启用状态

4. 刷新微信读书页面
   - 打开或刷新[微信读书](https://weread.qq.com/)页面，查看Word风格界面

## 更新日志

### v1.4.3 (最新版本)
- **最终版脚本 weread-word-style-final-v1.4.3.user.js**，功能和界面最完善
- 顶部蓝色条细节优化，增加自动保存、保存、撤销、重做、触摸/鼠标模式图标，调整搜索框和窗口控制按钮样式
- 所有功能区图标均为SVG，细节更贴近Word
- 多重初始化机制，兼容各种网络和浏览器环境
- 代码结构优化，提升加载和执行速度

### v1.4.1 ~ v1.3.1
- 详见脚本内注释和历史版本说明

## 注意事项

- 如安装了多个版本，请禁用旧版本，仅保留v1.4.3
- 如界面显示异常，请尝试清除浏览器缓存后刷新页面
- 本脚本仅修改界面样式，不收集任何用户数据，也不影响微信读书功能

## 问题反馈

如有问题或建议，请通过GitHub Issue或邮箱反馈。

## 许可证

本脚本基于MIT许可证开源。
