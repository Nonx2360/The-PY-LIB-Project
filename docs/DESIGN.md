# THE PY-LIB — Design System: "ห้องอ่านหนังสือ" (Reading Room)

One visual identity shared by the desktop app (CustomTkinter, `ui/theme.py`) and
the web console (`web/static/style.css`). Grounded in this project's actual
subject: a Thai school library, Thai-language staff UI, circulation-desk
workflows, and paper artifacts (member cards, due slips, PDF reports).

## Where the identity comes from

- **Navy** — sampled from `assets/logos/school_logo.png` (the school's own mark).
- **Paper neutrals** — the stationery of a library: index cards, due slips,
  member cards. Warm near-white, not cool slate.
- **Stamp status** — Thai school libraries still stamp due-date cards. Status
  badges are styled as rubber stamps: 2px ink border, tinted wash, 4px radius.
- **Sarabun body** — the Thai government's official school-document font; the
  app already bundles it for Thai text rendering. Prompt (Cadson Demak) sits
  above it for display.

## Tokens

| Token | Hex / value | Desktop source | Web source |
|---|---|---|---|
| paper (page bg) | `#F6F5F1` | `theme.CONTENT_LIGHT` | `--paper` |
| surface | `#FFFFFF` | `theme.CARD_LIGHT` | `--surface` |
| navy (primary) | `#1E4B63` | `theme.PRIMARY` | `--navy` |
| navy deep (sidebar) | `#14374A` | `theme.SIDEBAR` | `--navy-deep` |
| ink (text) | `#20303B` | `theme.TEXT_LIGHT` | `--ink` |
| ink secondary | `#4E6472` | `theme.MUTED_LIGHT` | `--ink-2` |
| marigold (accent) | `#B45309` / `#F59E0B` | `theme.WARNING` / ` WARNING_BRIGHT` | `--marigold` |
| stamp green | `#276749` | `theme.SUCCESS` | `--stamp-green` |
| stamp red | `#B3261E` | `theme.DANGER` | `--stamp-red` |
| borders | `#E3E0D8` (light) / `#2A3A46` (dark) | `theme.border_color()` | `--line` |

## Typography

- **Prompt SemiBold 600** — page titles, card titles, stat numbers, dashboard
  tile titles, brand wordmark. Bundled: `assets/fonts/Prompt-SemiBold.ttf`
  (FR_PRIVATE, registered in `theme.register_fonts()`, `DISPLAY_OK` flag).
- **Prompt Medium 500** — buttons.
- **Sarabun 400/600/700** — body, labels, nav items, tables, stamps.
- Thai rules: body line-height ≥ 1.6 (Thai ascenders/descenders clip at web
  defaults); never use uppercase/letter-spacing tricks (meaningless in Thai);
  line lengths under ~80 characters.

## Color discipline

- Navy for structure and primary actions. **Marigold is the single accent** and
  is spent only where attention is earned: active-nav notch, keyboard focus
  ring, the one highlighted element per screen. Never sprinkle it.
- Red is reserved for **overdue/failure only** ("ค้างส่ง", errors). Green for
  success/available. A stat row is all navy *except* the number that needs
  action — on dashboards, ค้างส่ง comes first and is the only red number.
- No decorative gradients, no rainbow icon colors, no hover translateY/scale
  (hover changes color/border only), no ALL-CAPS eyebrows, no "→" glyphs.

## Motion

Color transitions ≤ 150ms, user-triggered feedback only; the web honors
`prefers-reduced-motion` (`style.css`). Desktop uses no timed motion beyond
toast auto-dismiss.
