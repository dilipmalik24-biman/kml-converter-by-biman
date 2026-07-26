[app]
# App Title & Package
title = GIS Multi-Layer Converter
package.name = gismultilayer
package.domain = com.bimanmalik.gis

# Source code location
source.dir = .
source.include_exts = py,png,jpg,kv,atlas

# Application version
version = 1.0

# Dependencies (Exact versions to avoid build crashes)
requirements = python3,kivy==2.2.1,kivymd==1.1.1,openpyxl,pillow

# Orientation & Permissions
orientation = portrait
fullscreen = 0
android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE, MANAGE_EXTERNAL_STORAGE

# Android Specific Configuration
android.api = 33
android.minapi = 21
android.ndk = 25b
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_licenses = True
