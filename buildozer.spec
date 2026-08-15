[app]
title = Sauter Nadzor
package.name = sauternadzor
package.domain = org.sauter.nadzor
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,ttf
version = 1.0.0

requirements = python3,kivy,pyserial,openpyxl,opencv-python,pyzbar

orientation = landscape
fullscreen = 1

android.permissions = CAMERA,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,USB_PERMISSION
android.api = 33
android.minapi = 21
android.ndk = 25b
android.archs = arm64-v8a, armeabi-v7a

android.wakelock = True

[buildozer]
log_level = 2
warn_on_root = 1
