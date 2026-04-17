# FunctionBox

FunctionBox 是一个基于 VSTO 的 Word 加载项，面向日常文档核对和 VBA 工具管理场景。

## 主要功能

- 表格加总检查
  - 横向加总检查
  - 竖向加总检查
  - 非连续单元格加总检查
- 高亮清理
  - 清除选中高亮
  - 清除全文高亮
- VBA 工具箱
  - 保存常用 VBA 代码
  - 面板下拉执行 VBA
  - 录入并同步快捷键
  - 自动同步 Word 快捷键绑定
- 更新功能
  - 从 GitHub Release 检查新版本
  - 下载发布包并启动安装程序

## 项目结构

- [FunctionBox.sln](</E:/files/CodeProject/VS Project/FunctionBox/FunctionBox.sln>)
  - 解决方案入口
- [FunctionBox.csproj](</E:/files/CodeProject/VS Project/FunctionBox/FunctionBox.csproj>)
  - VSTO 项目入口
- [Ribbon1.cs](</E:/files/CodeProject/VS Project/FunctionBox/Ribbon1.cs>)
  - Ribbon 按钮事件入口
- [ThisAddIn.cs](</E:/files/CodeProject/VS Project/FunctionBox/ThisAddIn.cs>)
  - 加载项核心逻辑
- [Forms](</E:/files/CodeProject/VS Project/FunctionBox/Forms>)
  - 窗体与工具箱界面
- [Updater](</E:/files/CodeProject/VS Project/FunctionBox/Updater>)
  - 更新检查与发布包处理逻辑
- [Resources](</E:/files/CodeProject/VS Project/FunctionBox/Resources>)
  - 图标与图片资源
- [Properties](</E:/files/CodeProject/VS Project/FunctionBox/Properties>)
  - 程序集信息、资源和设置
- [.github/workflows](</E:/files/CodeProject/VS Project/FunctionBox/.github/workflows>)
  - GitHub Actions 构建与发布流程

## 开发环境

- Visual Studio 2026
- .NET Framework 4.8.1
- Microsoft Word / VSTO Runtime
- Windows

## 本地构建

```powershell
msbuild FunctionBox.sln /t:Build /p:Configuration=Debug /p:Platform="Any CPU"
```

## 发布与更新

- 版本号当前从 `AssemblyInfo.cs` 和 `FunctionBox.csproj` 统一维护
- GitHub Actions 在推送 `v*` 标签时构建并发布 Release
- 更新器从 GitHub Release 获取最新版本与发布包
- 发布包建议保持为 `FunctionBox.zip`
  - zip 内应包含 `setup.exe`

## CI 证书说明

仓库不需要提交 `FunctionBox_TemporaryKey.pfx`。

推荐做法：

- 将证书转为 Base64 后放入 GitHub Secrets
- 将证书密码放入 GitHub Secrets
- 在 Actions 中临时还原为 `.pfx`
- 通过 `CI_MANIFEST_KEY_FILE` 和 `CI_MANIFEST_KEY_PASSWORD` 传给 MSBuild

当前 workflow 已按这种方式配置。

## 备注

- 项目源码统一按 `UTF-8 with BOM` 保存，避免 VS 中中文乱码
- `bin`、`obj`、`.vs`、`.csproj.user`、本地证书文件等不应提交到 GitHub
