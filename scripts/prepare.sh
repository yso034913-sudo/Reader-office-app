#!/usr/bin/env bash
# Copies the web app into www/ and bundles every library locally so the APK works fully offline.
set -euo pipefail
rm -rf www && mkdir -p www/lib
cp index.html app.js ui.js office.js pptedit.js scan.js office.css style.css manifest.json icon-192.png icon-512.png icon-maskable-512.png native/native.js www/
B=https://cdnjs.cloudflare.com/ajax/libs
# ---- Office libs (from npm) — a missing one only disables that feature; office.js falls back to a CDN ----
cpf(){ # cpf <dest-name> <pkg-dir> <find-pattern>...
  local dest=$1 pkg=$2; shift 2; local f=""
  for pat in "$@"; do f=$(find "node_modules/$pkg" -type f -name "$pat" 2>/dev/null | head -n1 || true); [ -n "$f" ] && break; done
  if [ -n "$f" ]; then cp "$f" "www/lib/$dest"; echo "lib ok: $dest <- $f"; else echo "::warning::lib missing: $dest ($pkg)"; fi
}
cpf mammoth.browser.min.js mammoth mammoth.browser.min.js
cpf docx-preview.min.js docx-preview docx-preview.min.js
for f in node_modules/docx/build/index.umd.js node_modules/docx/dist/index.umd.cjs node_modules/docx/build/index.js; do
  if [ -f "$f" ]; then cp "$f" www/lib/docx.min.js; echo "lib ok: docx.min.js <- $f"; break; fi
done
[ -f www/lib/docx.min.js ] || echo "::warning::lib missing: docx.min.js"
cpf exceljs.min.js exceljs exceljs.min.js
# SheetJS 0.20.3 comes from package.json (official tarball); a missing copy is fatal because Excel needs it
cpf xlsx.full.min.js xlsx xlsx.full.min.js
[ -f www/lib/xlsx.full.min.js ] || { echo "::error::xlsx.full.min.js missing"; exit 1; }
cpf formula-parser.min.js hot-formula-parser formula-parser.min.js
cpf pptx-preview.umd.js pptx-preview pptx-preview.umd.js "*.umd.js"
cpf html2canvas.min.js html2canvas html2canvas.min.js
# pdf-lib: used by PowerPoint -> PDF export (pptedit.js loads lib/pdf-lib.min.js)
cpf pdf-lib.min.js pdf-lib pdf-lib.min.js
[ -f www/lib/pdf-lib.min.js ] || { echo "::error::pdf-lib.min.js missing (PowerPoint PDF export needs it offline)"; exit 1; }
# ---- ffmpeg.wasm (converts wma/aif/amr audio inside PowerPoint files to mp3) — optional, only that feature needs it ----
mkdir -p www/lib/ffmpeg
for pair in "@ffmpeg/ffmpeg/dist/umd/ffmpeg.js" "@ffmpeg/ffmpeg/dist/umd/814.ffmpeg.js" "@ffmpeg/core/dist/umd/ffmpeg-core.js" "@ffmpeg/core/dist/umd/ffmpeg-core.wasm"; do
  if [ -s "node_modules/$pair" ]; then
    cp "node_modules/$pair" "www/lib/ffmpeg/$(basename "$pair")"
    echo "lib ok: ffmpeg/$(basename "$pair")"
  else
    echo "::error::Required FFmpeg file missing: $pair"
    exit 1
  fi
done
# ESM core: needed when ffmpeg runs in a module worker (fallback path in office.js getFF)
mkdir -p www/lib/ffmpeg/esm
for f in ffmpeg-core.js ffmpeg-core.wasm; do
  if [ -s "node_modules/@ffmpeg/core/dist/esm/$f" ]; then cp "node_modules/@ffmpeg/core/dist/esm/$f" "www/lib/ffmpeg/esm/$f"; echo "lib ok: ffmpeg/esm/$f"; else echo "::warning::ESM core missing: $f"; fi
done
# the core libs are mandatory
cd www
# point the app at the local copies and drop the service worker (the APK already contains the files)
sed -i "s#https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/#lib/#g" index.html app.js
sed -i "s#^if('serviceWorker'in navigator.*register('sw.js').*;\$##" app.js
if grep -q "register('sw.js')" app.js; then echo "service worker removal failed"; exit 1; fi
sed -i 's#<script src="ui.js"></script>#<script src="ui.js"></script><script src="native.js"></script>#' index.html
grep -q 'native.js' index.html || { echo "native.js injection failed"; exit 1; }
du -sh .; echo "www ready"
