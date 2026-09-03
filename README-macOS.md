# KzzMonitor for macOS

## 系统要求

- macOS 13 或更新版本。
- Apple Silicon (`arm64`) 与 Intel (`x86_64`) 需要分别构建；不能混用。
- 运行 `.app` 不需要安装 Python。只有从源码构建时需要 Python 3.10+。

## 使用

1. 将整个 `release-macos` 文件夹复制到固定位置，例如 `~/Applications/KzzMonitor`。
2. 第一次启动时按住 Control 点击 `KzzMonitor.app`，选择“打开”，确认运行未签名应用。
3. 程序会在 `~/Library/Application Support/KzzMonitor/` 创建并使用：
   - `可转债监控.xlsx`
   - `data/monitor.db`
   - `logs/kzz_monitor.log`
4. 控制台中的“打开监控表”会打开实际使用的 Excel。
5. 关闭控制台窗口会隐藏至 macOS 菜单栏；菜单栏 `KZZ` 图标可恢复、启停、强制轮询或退出。

不要直接修改应用包旁边的模板工作簿；首次启动后应修改“打开监控表”打开的用户数据副本。

## SMTP 授权码

授权码存入 macOS 登录钥匙串，服务名为 `KzzMonitor SMTP`。macOS 可能在首次读取时询问是否允许访问，请选择允许。授权码不会写入 Excel。

## 通知权限

首次发送通知后，在“系统设置 → 通知”中允许 KzzMonitor/脚本通知。未签名的本地构建可能以 Python 或 Script Editor 名称出现。

## 登录与定时启动

在终端中运行：

    chmod +x install_macos_startup.sh uninstall_macos_startup.sh
    ./install_macos_startup.sh

它会创建当前用户的 LaunchAgent，在登录时和每天 09:20 启动。程序内部仍会判断交易日及开闭市时间。

取消：

    ./uninstall_macos_startup.sh

移动应用目录后，需要先取消再重新安装自启动。

## 从源码构建

在 Mac 上安装 Python 3.10+ 后：

    chmod +x build_macos.sh
    ./build_macos.sh

产物位于 `release-macos/KzzMonitor.app`。PyInstaller 不支持在 Windows 上交叉生成 macOS 应用。

## Gatekeeper、签名与分发

当前脚本生成的是本地未签名应用，适合个人使用。向其他 Mac 分发时，Gatekeeper 可能阻止启动。正式分发需要 Apple Developer ID，对 `.app` 执行 hardened runtime 签名、公证和 staple；这些步骤需要开发者证书和 Apple 凭据，本项目不会保存这些凭据。
