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
    for (Uri uri : uris) readAndSend(uri);
  }

  private void readAndSend(Uri uri) {
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
      JSONObject payload = new JSONObject();
      payload.put("name", name); payload.put("type", mime);
      payload.put("data", Base64.encodeToString(out.toByteArray(), Base64.NO_WRAP));
      String js = "window.dispatchEvent(new CustomEvent('readerall-file-received',{detail:"+payload.toString()+"}));";
      WebView web = getBridge().getWebView();
      if (web != null) web.postDelayed(new Runnable() {
        int tries=0;
        @Override public void run() {
          if (web.getUrl()!=null && web.getUrl().startsWith("http")) web.evaluateJavascript(js, null);
          else if (tries++<30) web.postDelayed(this, 300);
        }
      }, 900);
    } catch (Exception e) { android.util.Log.e("ReaderAll", "Could not receive shared file", e); }
  }
}
