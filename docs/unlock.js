/*
 * unlock.js — client-side decryption gate for the Hog Ranch dashboard.
 *
 * Loads BEFORE autoload.js and patches window.fetch. When autoload.js asks for
 * Hog_Ranch.xlsx / drill_pads.geojson / tenement.geojson, this intercepts the
 * request, fetches the .enc blob instead, decrypts it with a key derived from
 * the viewer's password, and hands back plaintext.
 *
 * autoload.js, manifest.json and index.html need no knowledge of any of this.
 *
 * Crypto: PBKDF2-HMAC-SHA256 (600k iterations) -> AES-256-GCM.
 * Blob layout: salt[16] || iv[12] || ciphertext+tag.
 * The key lives in memory only. Refreshing the page asks again.
 */
(function () {
  'use strict';

  var ENCRYPTED = ['Hog_Ranch.xlsx', 'drill_pads.geojson', 'tenement.geojson'];
  var PROBE = 'tenement.geojson'; // smallest file, used to validate the password
  var ITERATIONS = 600000;
  var SALT_LEN = 16;
  var IV_LEN = 12;

  var TYPES = {
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.geojson': 'application/geo+json',
    '.json': 'application/json'
  };

  if (!window.crypto || !window.crypto.subtle) {
    document.addEventListener('DOMContentLoaded', function () {
      document.body.innerHTML =
        '<p style="font:16px system-ui;padding:40px">This dashboard requires a ' +
        'secure (https) connection and a modern browser.</p>';
    });
    return;
  }

  /* ---------------------------------------------------------------- crypto */

  var keyResolve;
  var keyReady = new Promise(function (r) { keyResolve = r; });

  function deriveKey(password, salt) {
    return crypto.subtle
      .importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey'])
      .then(function (base) {
        return crypto.subtle.deriveKey(
          { name: 'PBKDF2', salt: salt, iterations: ITERATIONS, hash: 'SHA-256' },
          base,
          { name: 'AES-GCM', length: 256 },
          false,
          ['decrypt']
        );
      });
  }

  // Each blob carries its own salt, so a key is derived per file.
  function decryptBlob(buffer, password) {
    var bytes = new Uint8Array(buffer);
    if (bytes.length < SALT_LEN + IV_LEN + 16) {
      return Promise.reject(new Error('blob too short'));
    }
    var salt = bytes.slice(0, SALT_LEN);
    var iv = bytes.slice(SALT_LEN, SALT_LEN + IV_LEN);
    var body = bytes.slice(SALT_LEN + IV_LEN);
    return deriveKey(password, salt).then(function (key) {
      return crypto.subtle.decrypt({ name: 'AES-GCM', iv: iv }, key, body);
    });
  }

  /* -------------------------------------------------------- fetch shim */

  function basename(u) {
    return u.pathname.split('/').pop();
  }

  function encUrl(u) {
    var parts = u.pathname.split('/');
    parts[parts.length - 1] = parts[parts.length - 1] + '.enc';
    var out = new URL(u.href);
    out.pathname = parts.join('/');
    return out.href;
  }

  function contentTypeFor(name) {
    for (var ext in TYPES) {
      if (name.slice(-ext.length) === ext) return TYPES[ext];
    }
    return 'application/octet-stream';
  }

  var nativeFetch = window.fetch.bind(window);

  window.fetch = function (input, init) {
    var raw = typeof input === 'string' ? input : (input && input.url) || String(input);
    var url;
    try {
      url = new URL(raw, document.baseURI);
    } catch (e) {
      return nativeFetch(input, init);
    }

    var name = basename(url);
    if (ENCRYPTED.indexOf(name) === -1) return nativeFetch(input, init);

    // Blocks until the viewer has entered a working password.
    return keyReady.then(function (password) {
      return nativeFetch(encUrl(url), init).then(function (res) {
        // Pass real HTTP failures through so autoload.js reports them as usual.
        if (!res.ok) return res;
        return res.arrayBuffer().then(function (buf) {
          return decryptBlob(buf, password).then(function (plain) {
            return new Response(plain, {
              status: 200,
              statusText: 'OK',
              headers: { 'Content-Type': contentTypeFor(name) }
            });
          });
        });
      });
    });
  };

  /* ------------------------------------------------------------- overlay */

  var CSS =
    '#hr-lock{position:fixed;inset:0;z-index:2147483647;display:grid;place-items:center;' +
    'background:#10151c;font:15px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:#e6edf3;padding:24px}' +
    '#hr-lock .box{width:100%;max-width:360px;background:#171e27;border:1px solid #262f3b;' +
    'border-radius:12px;padding:32px 28px;box-shadow:0 12px 32px rgba(0,0,0,.4)}' +
    '#hr-lock h2{margin:0 0 4px;font-size:19px;font-weight:600}' +
    '#hr-lock .sub{margin:0 0 24px;font-size:13px;color:#8b98a8}' +
    '#hr-lock label{display:block;font-size:12px;font-weight:600;text-transform:uppercase;' +
    'letter-spacing:.06em;color:#8b98a8;margin-bottom:6px}' +
    '#hr-lock input{width:100%;box-sizing:border-box;padding:11px 12px;font-size:15px;' +
    'background:#0d1218;color:#e6edf3;border:1px solid #2d3846;border-radius:7px}' +
    '#hr-lock input:focus{outline:none;border-color:#4b93d1;box-shadow:0 0 0 3px rgba(75,147,209,.18)}' +
    '#hr-lock button{width:100%;margin-top:16px;padding:11px;font-size:15px;font-weight:600;' +
    'background:#2f7dc4;color:#fff;border:0;border-radius:7px;cursor:pointer}' +
    '#hr-lock button:disabled{opacity:.55;cursor:default}' +
    '#hr-lock .err{margin-top:14px;font-size:13px;color:#ff8a80;min-height:18px}' +
    '#hr-lock .note{margin-top:22px;font-size:11px;color:#5d6875;text-align:center}';

  function buildOverlay() {
    var style = document.createElement('style');
    style.textContent = CSS;
    document.head.appendChild(style);

    var el = document.createElement('div');
    el.id = 'hr-lock';
    el.innerHTML =
      '<div class="box" role="dialog" aria-modal="true" aria-label="Password required">' +
      '<h2>Hog Ranch Interactive Map</h2>' +
      '<p class="sub">Active Drilling Dashboard &middot; authorised access only</p>' +
      '<form id="hr-lock-form" autocomplete="off">' +
      '<label for="hr-lock-pw">Access password</label>' +
      '<input id="hr-lock-pw" type="password" required autocomplete="current-password">' +
      '<button id="hr-lock-go" type="submit">Unlock dashboard</button>' +
      '<div class="err" id="hr-lock-err" role="alert" aria-live="polite"></div>' +
      '</form>' +
      '<p class="note">Do not share this password outside your organisation.</p>' +
      '</div>';
    document.body.appendChild(el);

    var form = el.querySelector('#hr-lock-form');
    var input = el.querySelector('#hr-lock-pw');
    var btn = el.querySelector('#hr-lock-go');
    var err = el.querySelector('#hr-lock-err');

    input.focus();

    form.addEventListener('submit', function (e) {
      e.preventDefault();
      err.textContent = '';
      btn.disabled = true;
      btn.textContent = 'Unlocking…';

      var candidate = input.value;

      // Validate against the smallest encrypted file before letting anything
      // else through. Key derivation is deliberately slow (~0.3-1s).
      var probe = new URL(PROBE, document.baseURI);
      nativeFetch(encUrl(probe), { cache: 'no-cache' })
        .then(function (res) {
          if (!res.ok) throw new Error('missing:' + res.status);
          return res.arrayBuffer();
        })
        .then(function (buf) { return decryptBlob(buf, candidate); })
        .then(function () {
          el.remove();
          keyResolve(candidate); // releases every queued fetch
        })
        .catch(function (e2) {
          var missing = String(e2.message || '').indexOf('missing:') === 0;
          err.textContent = missing
            ? 'Encrypted data files are not published yet. Contact the site owner.'
            : 'Incorrect password.';
          input.value = '';
          input.focus();
          btn.disabled = false;
          btn.textContent = 'Unlock dashboard';
        });
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', buildOverlay);
  } else {
    buildOverlay();
  }
})();
