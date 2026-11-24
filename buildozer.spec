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

# 依赖：固定 Python 3.12（稳定版）+ Kivy + 安卓 Java 调用（用于导出文件到下载目录）
# 注意：p4a develop 默认 python3=3.14.2，Kivy 移动端对其支持不成熟易闪退；
#       python3 与 hostpython3 必须同时指定相同版本，否则 buildozer 会报错。
#       3.12 是 Kivy 2.3.1 长期验证的稳定组合，同时兼容较新的 NDK 29。
requirements = python3==3.12.10,hostpython3==3.12.10,kivy,pyjnius

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

# 支持主流架构：64 位 + 32 位
android.arch = arm64-v8a
android.archs = arm64-v8a, armeabi-v7a

android.accept_sdk_license = True
android.allow_backup = True
android.entrypoint = org.kivy.android.PythonActivity

# 关键修复：p4a 官方 master 分支自 2026-08 起损坏（capture_logging / charset_normalizer 构建错误），
# 修复已合入 develop 分支但尚未发版。显式指定 develop 才能正常构建。
p4a.branch = develop

[buildozer]
log_level = 2
warn_on_root = 1
