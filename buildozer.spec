[app]

# Application name
title = WhatsApp Chat Exporter

# Package name
package.name = whatsappexporter

# Package domain
package.domain = org.nabeen

# Source folder
source.dir = .

# Python files and other resources
source.include_exts = py,png,jpg,jpeg,kv,atlas,txt

# Application version
version = 1.0

# Python/Kivy dependencies
requirements = python3,kivy,openpyxl,pillow

# Screen orientation
orientation = portrait

# Fullscreen
fullscreen = 0


# ------------------------------------------------------------
# ANDROID
# ------------------------------------------------------------

# Permissions needed for your current file-based design
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# Architecture
android.archs = arm64-v8a

# Android app theme
android.entrypoint = org.kivy.android.PythonActivity


# ------------------------------------------------------------
# BUILD SETTINGS
# ------------------------------------------------------------

# Don't automatically add unnecessary files
exclude_exts = pyc,pyo

# Keep Python stdout/stderr for debugging
log_level = 2


[buildozer]

# Build warning level
log_level = 2

# Warn if running as root
warn_on_root = 1