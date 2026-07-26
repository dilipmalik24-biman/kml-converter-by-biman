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

# Dependencies
requirements = python3,kivy==2.2.1,kivymd,openpyxl,pillow

# Orientation & Permissions
orientation = portrait
fullscreen = 0
android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE, MANAGE_EXTERNAL_STORAGE

# Android Specific Configuration
android.api = 33
android.minapi = 21
android.ndk = 25b
android.archs = arm64-v8a, armeabi-v7a
