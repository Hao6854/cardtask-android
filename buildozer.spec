[app]
# 应用名称（安装到手机后显示的名字）
title = 储值卡与任务管理

# 包名（唯一标识，反域名写法）
package.name = cardtask
package.domain = com.workbuddy.tools

# 源码目录与需要打包进去的文件类型
source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,json
source.exclude_dirs = bin,build,.github

# 版本号（对应需求中的 1.0）
version = 1.0

# 依赖：python3 + Kivy + 安卓 Java 调用（用于导出文件到下载目录）
requirements = python3,kivy,pyjnius

# 竖屏
orientation = portrait
fullscreen = 0

# 不需要任何系统权限（数据存应用私有目录，导出走 MediaStore）
android.permissions =

# 兼容 Android 10 (API 29) 及以上
android.minapi = 29
android.api = 36
android.sdk = 36
android.ndk = 29
p4a.branch = develop

# 支持主流架构：64 位 + 32 位
android.arch = arm64-v8a
android.archs = arm64-v8a, armeabi-v7a

android.accept_sdk_license = True
android.allow_backup = True
android.entrypoint = org.kivy.android.PythonActivity

# 使用 Buildozer 自带的 python-for-android 默认版本，稳定性更好

[buildozer]
log_level = 2
warn_on_root = 1
