#!/usr/bin/env python3
import os, pathlib, xml.etree.ElementTree as ET
root=pathlib.Path('android/app/src/main')
manifest=root/'AndroidManifest.xml'
if not manifest.exists(): raise SystemExit('AndroidManifest.xml not found after Capacitor setup')
ET.register_namespace('android','http://schemas.android.com/apk/res/android')
A='{http://schemas.android.com/apk/res/android}'
t=ET.parse(manifest); m=t.getroot(); app=m.find('application')
# Camera and microphone permissions are needed only when the user starts scanning or recording audio.
if not any(x.get(A+'name')=='android.permission.CAMERA' for x in m.findall('uses-permission')):
  ET.SubElement(m,'uses-permission',{A+'name':'android.permission.CAMERA'})
if not any(x.get(A+'name')=='android.permission.RECORD_AUDIO' for x in m.findall('uses-permission')):
  ET.SubElement(m,'uses-permission',{A+'name':'android.permission.RECORD_AUDIO'})
# WebView getUserMedia needs MODIFY_AUDIO_SETTINGS too, otherwise it is denied even after RECORD_AUDIO is granted.
for perm in ('android.permission.MODIFY_AUDIO_SETTINGS',):
  if not any(x.get(A+'name')==perm for x in m.findall('uses-permission')):
    ET.SubElement(m,'uses-permission',{A+'name':perm})
if not any(x.get(A+'name')=='android.hardware.microphone' for x in m.findall('uses-feature')):
  ET.SubElement(m,'uses-feature',{A+'name':'android.hardware.microphone',A+'required':'false'})
# Android 11+: the app must declare that it queries the speech-recognition service.
q=m.find('queries')
if q is None: q=ET.SubElement(m,'queries')
if not any(a.get(A+'name')=='android.speech.RecognitionService' for i in q.findall('intent') for a in i.findall('action')):
  qi=ET.SubElement(q,'intent'); ET.SubElement(qi,'action',{A+'name':'android.speech.RecognitionService'})
if app is None: raise SystemExit('application element missing')
act=next((x for x in app.findall('activity') if x.get(A+'name','').endswith('MainActivity')),None)
if act is None: raise SystemExit('Capacitor MainActivity not found')
act.set(A+'exported','true')
# Separate filters keep Android MIME matching predictable for open/share.
mimes=[
'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
'application/vnd.ms-excel','text/csv',
'application/vnd.openxmlformats-officedocument.presentationml.presentation',
]
for action in ('android.intent.action.VIEW','android.intent.action.SEND','android.intent.action.SEND_MULTIPLE'):
  for mime in mimes:
    filt=ET.SubElement(act,'intent-filter')
    ET.SubElement(filt,'action',{A+'name':action})
    ET.SubElement(filt,'category',{A+'name':'android.intent.category.DEFAULT'})
    if action.endswith('VIEW'):
      ET.SubElement(filt,'category',{A+'name':'android.intent.category.BROWSABLE'})
      ET.SubElement(filt,'data',{A+'mimeType':mime,A+'scheme':'content'})
      ET.SubElement(filt,'data',{A+'mimeType':mime,A+'scheme':'file'})
    else:
      ET.SubElement(filt,'data',{A+'mimeType':mime})
t.write(manifest,encoding='utf-8',xml_declaration=True)
# Replace generated launcher activity with the receiver-enabled Capacitor activity.
pkg=pathlib.Path('android/app/src/main/java/com/readerall/app'); pkg.mkdir(parents=True,exist_ok=True)
src=pathlib.Path('native/MainActivity.java')
(pkg/'MainActivity.java').write_text(src.read_text())
print('Android intent filters and incoming-file activity installed')
