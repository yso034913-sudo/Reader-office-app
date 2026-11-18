#!/usr/bin/env bash
# Builds a debug APK: web files -> www/ -> Capacitor Android project -> Gradle.
set -euo pipefail
bash scripts/prepare.sh
[ -d android ] || npx cap add android
npx cap sync android
python3 scripts/patch-android.py
# app icon: legacy square icons + adaptive foreground/background
R=android/app/src/main/res
for d in mdpi hdpi xhdpi xxhdpi xxxhdpi; do
  mkdir -p $R/mipmap-$d
  cp icon-192.png $R/mipmap-$d/ic_launcher.png
  cp icon-192.png $R/mipmap-$d/ic_launcher_round.png
  cp native/ic_fg.png $R/mipmap-$d/ic_launcher_foreground.png
done
mkdir -p $R/values
printf '<?xml version="1.0" encoding="utf-8"?>\n<resources>\n <color name="ic_launcher_background">#0F172A</color>\n</resources>\n' > $R/values/ic_launcher_background.xml
cd android
chmod +x gradlew
./gradlew assembleDebug --no-daemon --stacktrace
# Copy the generated APK to the project root using the project name.
cp app/build/outputs/apk/debug/app-debug.apk "../Reader All.apk"
echo "APK ready: Reader All.apk"
