# Lawrence Kollel – Donor Lookup

Password-protected donor lookup site. Search donors by name/email/phone, filter gifts by
English or Hebrew calendar month, and open a profile for each donor (contact info,
pledges, balance, gift history). Hosted on GitHub Pages from the `docs/` folder.

The donor data is stored **encrypted** (`docs/data.enc.json`, AES-256-GCM, key derived
from the password with PBKDF2). The spreadsheet itself is never committed.

## Updating the data

1. Download the latest `2026 Donations.xlsx` into this folder.
2. Run (one-time: `pip install openpyxl cryptography`):

   ```
   python3 build.py "2026 Donations.xlsx" --password "THE_SITE_PASSWORD"
   ```

3. Commit and push:

   ```
   git add docs/data.enc.json
   git commit -m "Update donor data"
   git push
   ```

The site updates about a minute after the push.

## Changing the password

Re-run `build.py` with the new `--password` and push. The old password stops working
immediately (it can no longer decrypt the data).

## What the build reads

- `Sheet1` – contact info (title, phone, email, address)
- `Contribution 2026` – pledges, monthly donations (amount + method), and the Parlor Meeting block

Gifts are dated to the **month** they appear in on the sheet (the sheet has no day-level dates),
so the Hebrew-calendar view shows the one or two Hebrew months that each English month spans.
