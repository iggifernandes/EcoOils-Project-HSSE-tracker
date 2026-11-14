#!/usr/bin/env python3
"""Build the protected tracker (sign-in page + the tracker encrypted with gzip + AES-256-GCM).

  build/app.html        the real tracker, readable  (NOT committed: .gitignore)
  build/secrets/app.key 256-bit key, base64         (NOT committed). The same key is APP_KEY_B64 in the deployed Code.gs.
  build/shell.html      the sign-in page template (committed)

  python3 build/build.py             -> next/index.html   (staging, safe to test)
  python3 build/build.py --live      -> index.html        (what everybody opens)
"""
import base64, gzip, os, re, sys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

here = os.path.dirname(os.path.abspath(__file__))
root = os.path.dirname(here)
rd = lambda p: open(os.path.join(here, p), encoding="utf-8").read()

app = rd("app.html")
key = base64.b64decode(rd("secrets/app.key").strip())
assert len(key) == 32, "app.key must be 32 bytes (base64)"
assert "viewToken" in app, "app.html does not carry the sign-in token - wrong file?"
ver = re.search(r'const DISPLAY_VERSION = "([^"]+)"', app).group(1)
api = re.search(r'const GOOGLE_SHEET_API_URL = "([^"]+)"', app).group(1)
logo = base64.b64encode(open(os.path.join(here, "logo_small.png"), "rb").read()).decode()

packed = gzip.compress(app.encode("utf-8"), compresslevel=9, mtime=0)
iv = os.urandom(12)
payload = base64.b64encode(iv + AESGCM(key).encrypt(iv, packed, None)).decode()   # iv || ciphertext || tag (WebCrypto layout)

shell = rd("shell.html")
for k, v in (("__VER__", ver), ("__API__", api), ("__LOGO__", logo), ("__PAYLOAD__", payload)):
    assert k in shell, k
    shell = shell.replace(k, v)

dest = os.path.join(root, "index.html") if "--live" in sys.argv else os.path.join(root, "next", "index.html")
os.makedirs(os.path.dirname(dest), exist_ok=True)
open(dest, "w", encoding="utf-8").write(shell)
print("v%s  app %d KB -> gzip %d KB -> published %d KB  =>  %s" % (ver, len(app) // 1024, len(packed) // 1024, len(shell) // 1024, os.path.relpath(dest, root)))
