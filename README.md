# PandocGUI

基于 WinUI 3 的 Pandoc 图形界面，面向 Windows 10/11。主打“设置式”体验：导入、配置、队列与日志都可视化完成。

## 功能概览
- 多文件转换、队列化处理、进度与日志显示
- 支持 Pandoc 全部功能参数（自由追加参数）
- 预设管理（保存/应用常用参数组合）
- 输出格式专属设置（按格式覆盖扩展名、模板、额外参数）
- Pandoc 自动检测、手动选择、或从 GitHub 下载官方 Windows 版
- 最近使用记录（文件、输出目录、模板、输出格式）
- 可选“拖入即开始”、可配置最大并行度

## 运行环境
- Windows 10 1809+（TargetPlatformMinVersion 10.0.17763）
- .NET 9 SDK
- Visual Studio 2022 或 Build Tools（可选，命令行 build/publish 也可）

## 目录结构
```
PandocGUI/
├─ Assets/                 图标与资源
├─ Models/                 数据模型
├─ Services/               设置/检测/下载等服务
├─ ViewModels/             视图模型
├─ Views/                  页面
├─ Properties/PublishProfiles/
├─ build_x64.bat            清理 + 构建
├─ publish_x64_portable.bat 便携版发布
├─ publish_x64_msix.bat      MSIX 发布（不签名）
└─ README.md
```

## 主要页面
- 首页：应用概览与入口
- 转换：文件列表 + 右侧配置面板（输出、模板、参数、预设、队列状态）
- 输出格式：为指定输出格式设置扩展名/模板/额外参数
- 队列：进度、明细、日志、暂停/继续/清理
- 设置：Pandoc 检测、下载、并行度、自动开始等

## Pandoc 检测与下载
检测顺序（逐级命中即可）：
1. 当前选择的路径
2. 已保存路径（settings.json）
3. `where pandoc` 结果
4. 扫描 PATH 目录

检测过程会输出详细步骤与命中来源。

下载：可从 GitHub 拉取官方 Windows 版（windows-x86_64.zip），解压到：
```
%LocalAppData%\PandocGUI\pandoc
```
下载完成后会自动写入 Pandoc 路径到设置。

## 数据文件与设置
应用运行目录下会生成：
```
data/settings.json
```
包含：
- 首次启动标记
- Pandoc 路径
- 输出目录、输入/输出格式、模板、附加参数
- 预设列表
- 输出格式专属设置
- 最近使用记录
- 并行度、拖入即开始等

## 常用流程
1. 打开应用，首次检测/设置 Pandoc
2. 拖拽或选择文件加入队列
3. 配置输出格式、输出目录、模板或参数
4. 需要时保存为预设
5. 点击开始转换，查看进度与日志

## 构建
清理并构建（x64）：
```
build_x64.bat
```

## 发布
便携版（无需安装）：
```
publish_x64_portable.bat
```
输出目录：
```
bin\Release\net9.0-windows10.0.19041.0\win-x64\publish\
```

MSIX 安装包（不签名）：
```
publish_x64_msix.bat
```
输出目录：
```
AppPackages\
```

## 常见问题
**Q: 为什么检测不到 Pandoc？**  
A: 请在设置页手动选择 `pandoc.exe`，或确保 `pandoc` 已加入 PATH；也可直接使用“下载官方版本”。

**Q: 可以只打包 bin 吗？**  
A: 建议使用 `publish` 生成便携版，或使用 MSIX 生成安装包；`bin` 更偏向构建产物，不推荐直接分发。

**Q: 输出格式相关的配置在哪里？**  
A: 在“输出格式”页面为每个格式单独设置（扩展名/模板/额外参数），会自动应用到转换参数。

