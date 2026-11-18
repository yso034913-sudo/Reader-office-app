package com.readerall.app;

import android.content.Intent;
import android.content.ClipData;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.provider.OpenableColumns;
import android.util.Base64;
import android.webkit.WebView;
import com.getcapacitor.BridgeActivity;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import org.json.JSONObject;

public class MainActivity extends BridgeActivity {
  @Override public void onCreate(Bundle savedInstanceState) {
    super.onCreate(savedInstanceState);
    deliverIntent(getIntent());
  }

  @Override public void onNewIntent(Intent intent) {
    super.onNewIntent(intent);
    setIntent(intent);
    deliverIntent(intent);
  }

  private void deliverIntent(Intent intent) {
    if (intent == null) return;
    String action = intent.getAction();
    ArrayList<Uri> uris = new ArrayList<>();
    if (Intent.ACTION_SEND.equals(action)) {
      Uri u = intent.getParcelableExtra(Intent.EXTRA_STREAM);
      if (u == null && intent.getData() != null) u = intent.getData();
      if (u != null) uris.add(u);
    } else if (Intent.ACTION_SEND_MULTIPLE.equals(action)) {
      ClipData clip = intent.getClipData();
      if (clip != null) for (int i=0;i<clip.getItemCount();i++) if (clip.getItemAt(i).getUri()!=null) uris.add(clip.getItemAt(i).getUri());
      if (uris.isEmpty()) {
        try {
          ArrayList<Uri> list = intent.getParcelableArrayListExtra(Intent.EXTRA_STREAM);
          if (list != null) uris.addAll(list);
        } catch (Exception ignored) {}
      }
    } else if (Intent.ACTION_VIEW.equals(action) && intent.getData() != null) {
      uris.add(intent.getData());
    }
    final WebView web = getBridge() != null ? getBridge().getWebView() : null;
    if (web == null) return;
    for (final Uri uri : uris) new Thread(new Runnable() { @Override public void run() { readAndSend(uri, web); } }).start();
  }

  private static final int CHUNK = 262144; // chars of base64 per evaluateJavascript call (multiple of 4)

  // Runs on a worker thread: reads the file, then hands it to the page in chunks once the page is ready.
  private void readAndSend(Uri uri, final WebView web) {
    try {
      String name = "office-file";
      Cursor c = getContentResolver().query(uri, null, null, null, null);
      if (c != null) {
        try { int ix=c.getColumnIndex(OpenableColumns.DISPLAY_NAME); if(c.moveToFirst() && ix>=0) name=c.getString(ix); }
        finally { c.close(); }
      }
      String mime = getContentResolver().getType(uri);
      if (mime == null) mime = "application/octet-stream";
      ByteArrayOutputStream out = new ByteArrayOutputStream();
      try (InputStream in = getContentResolver().openInputStream(uri)) {
        if (in == null) return;
        byte[] buf = new byte[16384]; int n;
        while ((n=in.read(buf))>0) out.write(buf,0,n);
      }
      final String b64 = Base64.encodeToString(out.toByteArray(), Base64.NO_WRAP);
      out = null;
      final String id = java.util.UUID.randomUUID().toString();
      final String qName = JSONObject.quote(name), qMime = JSONObject.quote(mime), qId = JSONObject.quote(id);
      final int total = Math.max(1, (b64.length() + CHUNK - 1) / CHUNK);
      web.postDelayed(new Runnable() {
        int tries = 0;
        @Override public void run() {
          // wait until the page has loaded native.js (it sets window.__raReady)
          web.evaluateJavascript("window.__raReady===true", new android.webkit.ValueCallback<String>() {
            @Override public void onReceiveValue(String v) {
              if ("true".equals(v)) {
                for (int i = 0; i < total; i++) {
                  String part = b64.substring(i * CHUNK, Math.min(b64.length(), (i + 1) * CHUNK));
                  web.evaluateJavascript("window.__raChunk("+qId+","+qName+","+qMime+","+i+","+total+","+JSONObject.quote(part)+");", null);
                }
              } else if (tries++ < 100) {
                web.postDelayed(Runnable_this(), 300);
              } else {
                android.util.Log.e("ReaderAll", "Page never became ready; shared file dropped");
              }
            }
          });
        }
        private Runnable Runnable_this() { return this; }
      }, 500);
    } catch (Exception e) { android.util.Log.e("ReaderAll", "Could not receive shared file", e); }
  }
}
