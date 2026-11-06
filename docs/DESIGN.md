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

- **Prompt SemiBold 600** — brand wordmark, page titles, card titles, ledger
  numbers. Bundled: `assets/fonts/Prompt-SemiBold.ttf`
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
  success/available. A number row is all ink *except* the value that needs
  action — on dashboards, ค้างส่ง comes first and is the only red number.
- No decorative gradients, no rainbow icon colors, no hover translateY/scale
  (hover changes color/border only), no ALL-CAPS eyebrows, no "→" glyphs.

## Composition rules (web console)

These are fixed patterns in `web/static/style.css`, not per-page
improvisations — a new page reuses them before inventing its own:

- **Ledger strip** (`.ledger`) — the dashboard's four numbers render as one
  card with hairline dividers, display-font numerals and
  `font-variant-numeric: tabular-nums`. Never a row of identical stat tiles.
  Red only on ค้างส่ง, which leads.
- **Quick-link pills** (`.quick-links`) — dashboard shortcuts are compact
  pills: the two circulation-desk pages (ยืม / คืน) are solid navy, the
  management links stay outlined. Never large centered icon tiles.
- **Chart hierarchy** (`.chart-grid`) — the borrow trend is the hero
  (~`2fr`); access and status share a stacked `1fr` side column. Three equal
  chart boxes is exactly the flatness to avoid.
- **The card grammar — เอกสาร form look** — every module (`.card`) is a
  ruled school form: a `4px double` navy letterhead across the top, a
  tinted header band (`--paper-band`) with the title set in navy-deep ink,
  clean white paper below. This is the identity workhorse; it shows on
  every staff page and the member portal at once.
- **The date stamp** — a marigold stamp chip (`--date-stamp`) in the staff
  top bar and on the login slip: `วันพฤหัสบดี 6 พฤศจิกายน 2569`, computed
  server-side (`stamp_date` context processor). The desk's world is dates
  (due dates, access times); the stamp is its most native mark. Marigold
  is spent here + focus ring + active-nav notch, nowhere else.
- **Card headers are text** — `h3` carries no decorative icon. Icons are
  spent on actions only (buttons, scan rows, sidebar nav), where they mean
  something.
- **Scanner rows are desk slips** — `.scan-row` is a dashed, paper-tinted
  strip: the camera module reads as an intake step, not a random button.
- **Empty states** (`.empty-state`) — a quiet block with a title plus a
  one-line note; never blank space, never layout hacks (`line-height` on a
  spacer paragraph is banned).
- **Login** (`.login-slip`) — one paper slip with a navy top rule and the
  school line under a hairline. No split-panel marketing hero.
- **Login copy** states what the software is — "ระบบบริหารห้องสมุดโรงเรียน".
  No self-praise ("อัจฉริยะ").
- **Numerals** — tabular figures in tables and the ledger so counts align
  like a register column.

### Thai typography guardrails

- Thai has no case: **no `text-transform: uppercase` and no
  `letter-spacing` on Thai text** — spacing breaks combining marks
  (`สระ`, `-ั`, `-ี`). Latin wordmarks/codes may keep letter-spacing.
- Line-height stays ≥ 1.6 for body text (Thai ascenders/descenders clip at
  web defaults).

## Motion

Color transitions ≤ 150ms, user-triggered feedback only; the web honors
`prefers-reduced-motion` (`style.css`). Desktop uses no timed motion beyond
toast auto-dismiss.
