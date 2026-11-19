# 大文件免费版说明

本项目仓库：[zhishouxun/vscode-office](https://github.com/zhishouxun/vscode-office/tree/master)。

上游仓库：[cweijan/vscode-office](https://github.com/cweijan/vscode-office)。

本地扩展版本为 `4.2.3-local.3`，扩展 ID 为 `local.office-largefile-local`。改动包括大文件分页编辑、滚动自动翻页、可调整分页容量、全文搜索准确定位、阅读位置与光标恢复、相邻页预处理缓存，以及内置功能的本地免费使用策略。

优化版以官方发布的 4.2.3 扩展为基础。当前仓库的上游源码基线为 4.2.2，与已发布扩展存在差异，因此修改覆盖层单独保存在 `local-largefile/` 中。根目录的常规构建流程不会自动应用这些覆盖层；请按下面的说明重建本地版。

- [使用、构建与验证说明](local-largefile/README.md)
- [分页与全文搜索模型](local-largefile/resource/markdown/largeDocument.mjs)
- [分页编辑与阅读位置恢复](local-largefile/resource/markdown/largeDocumentUI.mjs)
- [文字导航辅助](local-largefile/resource/markdown/documentNavigation.mjs)
- [本地功能构建策略](local-largefile/local_features.py)
- [资源覆盖层补丁](local-largefile/upstream.patch)

在仓库根目录运行模型测试：

```sh
node --test local-largefile/test/largeDocument.test.mjs
```

构建后的 VSIX 位于指定构建目录，可使用 VS Code 的“从 VSIX 安装”功能安装。本地生成的 `dist/`、依赖目录和测试输出不纳入版本管理。
