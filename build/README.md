# Protected tracker - how it works

* `index.html` (what people open) = a small sign-in page + the real tracker **encrypted** (gzip + AES-256-GCM).
* Sign-in: the person types their @ecooils.com.my / @ecooils.co.id email, receives a 6-digit code (Apps Script `Code.gs`, MailApp),
  types it, and the server returns a signed session token (14 days) **and the decryption key**. The key is never stored in the page.
* Every code request, sign-in and page open is logged in the Google Sheet tab `AccessLog` (tab `AccessSummary` = per person).
  Viewers cannot see who else opened the tracker. Block someone: add their email to the `Blocked` tab. Sign everyone out: Sheet menu.
* The key is `build/secrets/app.key` (NOT in git) and the same value is `APP_KEY_B64` in the deployed Code.gs (blank in the repo copy).

## Changing the tracker
1. `python3 build/decrypt.py`  -> readable `build/app.html` (needs `build/secrets/app.key`)
2. edit `build/app.html` (bump DISPLAY_VERSION)
3. `python3 build/build.py` (staging: `next/index.html`) test, then `python3 build/build.py --live` (`index.html`) and push.
Never commit `build/app.html` or `build/secrets/` (they are in .gitignore).

If the key is lost: make a new one (`python3 -c "import os,base64;print(base64.b64encode(os.urandom(32)).decode())"`), put it in
`build/secrets/app.key` and in Code.gs `APP_KEY_B64`, redeploy Code.gs, rebuild. Rotate it if a person who knew it should lose access.
