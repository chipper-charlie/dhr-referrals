# Referrals finder prototype

A static test site for the NYDHR referrals finder. Not the live NYDHR website.

- `/` choose a version, `/public` the public finder, `/staff` the staff helper.
- Hosted on Vercel as a plain static site: no build step (Framework preset "Other").
- Kept out of search engines: `noindex` on every page, an `X-Robots-Tag` header (vercel.json) and robots.txt.
- `data/` is the finder data, copied from `Referrals/built/Referral_Database.xlsx` by `Referrals/scripts/export_finder_data.py`. To update the site, re-run the build and export, copy `Referral finder/data/` here, and push.
