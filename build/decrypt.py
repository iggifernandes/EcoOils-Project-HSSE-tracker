#!/usr/bin/env python3
"""Get the readable tracker back out of the published page (needs build/secrets/app.key).

  python3 build/decrypt.py [index.html]      -> build/app.html
"""
import base64, gzip, os, re, sys
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

here = os.path.dirname(os.path.abspath(__file__))
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(here), "index.html")
key = base64.b64decode(open(os.path.join(here, "secrets", "app.key"), encoding="utf-8").read().strip())
page = open(src, encoding="utf-8").read()
blob = base64.b64decode(re.search(r'<script id="p" type="text/plain">([^<]+)</script>', page).group(1))
plain = gzip.decompress(AESGCM(key).decrypt(blob[:12], blob[12:], None))
open(os.path.join(here, "app.html"), "wb").write(plain)
print("wrote build/app.html (%d KB)" % (len(plain) // 1024))
