[app]
title = Sauter Nadzor
package.name = sauternadzor
package.domain = org.sauter.nadzor
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,ttf
version = 1.0.0

requirements = python3==3.10.13,hostpython3==3.10.13,kivy,pyserial,openpyxl,et_xmlfile

orientation = landscape
fullscreen = 1

android.permissions = CAMERA,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.ndk = 25b
android.archs = arm64-v8a
android.allow_backup = True

android.wakelock = True
android.accept_sdk_license = True
android.skip_update = False

[buildozer]
log_level = 2
warn_on_root = 1
