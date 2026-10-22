/* An Tâm Số – cổng mật khẩu
 * Dữ liệu trên web chỉ có dạng mã hóa (data/data.enc.js: PBKDF2-SHA256 → AES-256-CBC + HMAC-SHA256).
 * Mã nguồn không chứa mật khẩu; chỉ người nhập đúng mật khẩu mới giải mã được số liệu.
 * Khi mở trên máy nội bộ sau update.bat, data/data.local.js (không đẩy lên GitHub) được dùng trực tiếp. */
(function () {
  'use strict';
  var KEY_STORE = 'ats_dash_key_v1';
  var te = new TextEncoder(), td = new TextDecoder();
  var b64 = function (s) { var bin = atob(s), u = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
  var tob64 = function (u) { var s = ''; u = new Uint8Array(u); for (var i = 0; i < u.length; i++) s += String.fromCharCode(u[i]); return btoa(s); };
  var store = {
    get: function () { try { return localStorage.getItem(KEY_STORE); } catch (e) { return null; } },
    set: function (v) { try { localStorage.setItem(KEY_STORE, v); } catch (e) {} },
    del: function () { try { localStorage.removeItem(KEY_STORE); } catch (e) {} }
  };

  function deriveKeys(password, E) {
    return crypto.subtle.importKey('raw', te.encode(password), 'PBKDF2', false, ['deriveBits'])
      .then(function (base) { return crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: te.encode(E.salt), iterations: E.iter }, base, 512); })
      .then(function (bits) { return new Uint8Array(bits); });
  }
  function decrypt(keys, E) {
    var encK = keys.slice(0, 32), macK = keys.slice(32, 64);
    var iv = b64(E.iv), ct = b64(E.ct), mac = b64(E.mac);
    var signed = new Uint8Array(iv.length + ct.length); signed.set(iv, 0); signed.set(ct, iv.length);
    return crypto.subtle.importKey('raw', macK, { name: 'HMAC', hash: 'SHA-256' }, false, ['verify'])
      .then(function (k) { return crypto.subtle.verify('HMAC', k, mac, signed); })
      .then(function (ok) {
        if (!ok) throw new Error('bad-password');
        return crypto.subtle.importKey('raw', encK, { name: 'AES-CBC' }, false, ['decrypt']);
      })
      .then(function (k) { return crypto.subtle.decrypt({ name: 'AES-CBC', iv: iv }, k, ct); })
      .then(function (pt) { return JSON.parse(td.decode(pt)); });
  }

  function lockUI(onSubmit) {
    var css = document.createElement('style');
    css.textContent = '\
#atsLock{position:fixed;inset:0;z-index:100;display:flex;align-items:center;justify-content:center;padding:20px;font-family:Inter,system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:#0b1e3f;\
background:radial-gradient(800px 500px at 100% 0%,rgba(29,99,216,.10),transparent 60%),radial-gradient(700px 500px at 0% 0%,rgba(225,29,46,.07),transparent 60%),radial-gradient(circle at 1px 1px,rgba(11,30,63,.08) 1px,transparent 0) 0 0/22px 22px,#f2f5f9}\
#atsLock .bx{position:relative;width:min(400px,100%);background:#fff;border:1px solid #e3e9f1;border-radius:20px;padding:30px 26px 24px;box-shadow:0 30px 70px -30px rgba(11,30,63,.35);overflow:hidden;animation:atsIn .4s ease}\
#atsLock .bx::after{content:"";position:absolute;left:0;right:0;bottom:0;height:3px;background:linear-gradient(90deg,#e11d2e,#0fa9a3,#1d63d8)}\
#atsLock .ic{width:64px;height:64px;border-radius:16px;overflow:hidden;margin:0 auto 14px;box-shadow:0 10px 24px -10px rgba(11,30,63,.5)}\
#atsLock .ic img{width:100%;height:100%;display:block}\
#atsLock h1{display:block;font-family:"Plus Jakarta Sans",Inter,sans-serif;font-size:21px;font-weight:800;text-align:center;letter-spacing:-.01em}\
#atsLock p{font-size:13px;color:#61718a;text-align:center;margin:6px 0 20px;line-height:1.5}\
#atsLock .tag{display:block;width:max-content;margin:0 auto 10px;font-family:"JetBrains Mono",ui-monospace,Consolas,monospace;font-size:10.5px;color:#0a7f7b;background:#e4f6f5;border-radius:999px;padding:4px 10px;letter-spacing:.06em}\
#atsLock .f{position:relative}\
#atsLock input[type=password],#atsLock input[type=text]{width:100%;font-size:16px;padding:13px 46px 13px 14px;border:1px solid #cfd9e6;border-radius:12px;outline:none;font-family:inherit;color:#0b1e3f;background:#f7f9fc}\
#atsLock input:focus{border-color:#0fa9a3;box-shadow:0 0 0 3px rgba(15,169,163,.15);background:#fff}\
#atsLock .eye{position:absolute;right:8px;top:50%;transform:translateY(-50%);border:0;background:transparent;cursor:pointer;color:#98a5b8;padding:6px;font-size:13px}\
#atsLock label{display:flex;align-items:center;gap:8px;font-size:13px;color:#61718a;margin:12px 0 16px;cursor:pointer;user-select:none}\
#atsLock label input{accent-color:#0fa9a3;width:16px;height:16px}\
#atsLock button.go{width:100%;font-size:15px;font-weight:700;padding:13px;border:0;border-radius:12px;background:#e11d2e;color:#fff;cursor:pointer;font-family:inherit;box-shadow:0 8px 20px -10px rgba(225,29,46,.8)}\
#atsLock button.go:disabled{opacity:.6;cursor:wait}\
#atsLock .err{min-height:18px;font-size:12.5px;color:#e11d2e;text-align:center;margin-top:10px;font-weight:600}\
#atsLock .ft{font-size:11px;color:#98a5b8;text-align:center;margin-top:8px}\
#atsLock.shake .bx{animation:atsShake .35s}\
@keyframes atsIn{from{opacity:0;transform:translateY(10px)}}\
@keyframes atsShake{25%{transform:translateX(-8px)}50%{transform:translateX(8px)}75%{transform:translateX(-4px)}}';
    document.head.appendChild(css);
    var el = document.createElement('div');
    el.id = 'atsLock';
    el.innerHTML = '<form class="bx" autocomplete="on"><div class="ic"><img src="icons/icon-192.png" alt=""></div>' +
      '<span class="tag">KHU VỰC NỘI BỘ · ENCRYPTED</span>' +
      '<h1>Bảo hiểm An Tâm Số</h1><p>Báo cáo doanh thu BH PVI × MISA.<br>Nhập mật khẩu để xem số liệu.</p>' +
      '<input type="text" name="username" value="antamso" autocomplete="username" style="display:none">' +
      '<div class="f"><input type="password" id="atsPw" placeholder="Mật khẩu" autocomplete="current-password" required><button type="button" class="eye" id="atsEye">Hiện</button></div>' +
      '<label><input type="checkbox" id="atsRem" checked>Ghi nhớ trên thiết bị này</label>' +
      '<button class="go" type="submit" id="atsGo">Mở khóa</button><div class="err" id="atsErr"></div>' +
      '<div class="ft">Dữ liệu được mã hóa AES-256 · chỉ giải mã trên thiết bị của bạn</div></form>';
    document.body.appendChild(el);
    var pw = el.querySelector('#atsPw'), go = el.querySelector('#atsGo'), err = el.querySelector('#atsErr');
    el.querySelector('#atsEye').onclick = function () { var t = pw.type === 'password'; pw.type = t ? 'text' : 'password'; this.textContent = t ? 'Ẩn' : 'Hiện'; };
    el.querySelector('form').onsubmit = function (e) {
      e.preventDefault(); if (!pw.value) return;
      go.disabled = true; go.textContent = 'Đang giải mã…'; err.textContent = '';
      onSubmit(pw.value, el.querySelector('#atsRem').checked).then(function () {
        el.style.transition = 'opacity .35s'; el.style.opacity = '0'; setTimeout(function () { el.remove(); }, 380);
      }, function () {
        go.disabled = false; go.textContent = 'Mở khóa'; err.textContent = 'Sai mật khẩu, vui lòng thử lại.';
        el.classList.remove('shake'); void el.offsetWidth; el.classList.add('shake'); pw.select();
      });
    };
    setTimeout(function () { pw.focus(); }, 50);
  }

  window.Gate = {
    lock: function () { store.del(); location.reload(); },
    run: function (start) {
      if (window.DASH_DATA) { start(); return; }
      var E = window.DASH_ENC;
      if (!E || !window.crypto || !crypto.subtle) { start(); return; }
      var ok = function (data) { window.DASH_DATA = data; window.__fromEnc = true; start(); };
      var saved = store.get();
      var tryPw = function (password, remember) {
        return deriveKeys(password, E).then(function (keys) {
          return decrypt(keys, E).then(function (data) { if (remember) store.set(tob64(keys)); ok(data); });
        });
      };
      if (saved) {
        decrypt(b64(saved), E).then(ok, function () { store.del(); lockUI(tryPw); });
      } else {
        lockUI(tryPw);
      }
    }
  };
})();
