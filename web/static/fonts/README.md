# vendored web fonts

Copied from `assets/fonts/` so the web UI loads zero CDN traffic (the
school LAN must work with no internet). Both families are released under
the SIL Open Font License — see `OFL-Prompt.txt`; Sarabun is licensed
under the same OFL (16 February 2012 text).

| File | Weight served | Used for |
|---|---|---|
| `Prompt-Medium.ttf` | 500 | `.action-text`, side headings |
| `Prompt-SemiBold.ttf` | 600–700 (weight range) | titles, headings, stat numbers |
| `Sarabun-Regular.ttf` | 400 | body text |
| `Sarabun-Bold.ttf` | 600–700 (weight range) | strong body text, labels |

The 600–700 weight range registration is deliberate: only four TTFs
exist on disk (the missing Sarabun SemiBold could not be fetched from
the CDN at vendoring time), so the Bold face stands in for both weights —
the browser serves the true glyphs instead of synthesizing bold.

These faces are referenced only by `@font-face` rules at the top of
`../style.css`; no template links to `fonts.googleapis.com` anymore.
