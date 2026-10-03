[app]

title = WhatsApp Chat Exporter
package.name = whatsappchatexporter
package.domain = org.nabeen

source.dir = .
source.include_exts = py,png,jpg,jpeg,gif,webp,bmp,kv,atlas,txt

requirements = python3,kivy,openpyxl,pillow,pyjnius

version = 1.0.0

orientation = portrait
fullscreen = 0

android.permissions = MANAGE_EXTERNAL_STORAGE,READ_MEDIA_IMAGES,READ_MEDIA_VIDEO,READ_MEDIA_AUDIO,POST_NOTIFICATIONS

android.api = 35
android.minapi = 23

android.archs = arm64-v8a

entrypoint = main.py

android.accept_sdk_license = True

log_level = 2
