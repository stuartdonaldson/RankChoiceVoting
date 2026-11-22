# UI design: RankChoiceVoting implementation notes

The design standards are owned by DevStandard, not this repo (`$DEVSTANDARD` = /home/stuar/proj/DevStandard):

| Doc | Owns |
|---|---|
| `ui-framework/01-ui-design-principles.md` | Outcomes, validation tiers (01 §4) |
| `ui-framework/02-shared-ui-component-standards.md` | Token contract and contrast values (§2.3), component contracts, header/identity (§4), acceptance criteria BR-1..BR-6, ID-1..ID-8, HD-1 (§10) |
| `ui-framework/03-northlake-design-language.md` | Northlake artwork, palette, typography, geometry |

This file keeps only what is specific to this app. The earlier full design system (draft 0.3, from
which `northlake.css` was built) is kept unchanged at `docs/archive/nuuc-ui-design-v0.3.md` for the
sections ui-framework does not carry: §2.5 logo deployment snippet, §10 icon mapping table, §14
theme-selection reference code, §17 governance. Where they disagree, ui-framework wins.

## Themes per env

`THEME` in `tools/static-pages.js`: `sit` = northlake, `nuuc` = northlake (rcballot-gk3, Communications
palette approved), `prod` = f3 (its own palette, fonts and icons; unchanged, and guarded by
`test/test_static_themes.js`). Each theme is one CSS module in `static-pages/src/themes/`, scoped to
`html[data-theme="<name>"]`; a module never names another theme. The stamped page links only its
env's module (`themeCssLink_`).

## Token mapping (northlake.css `--nl-*` -> 02 `--ui-*`)

The module keeps `--nl-*` names and maps them onto the page's app tokens (`--bg-page`, `--accent`,
...) in its first block. Values equal 02 §2.3 in both schemes; `test/test_static_themes.js` asserts them
and their contrast.

| `--nl-*` | 02 `--ui-*` |
|---|---|
| background / surface / surface-subtle | bg / surface / surface-muted |
| text / text-muted / heading | text / text-muted / heading |
| primary / on-primary / primary-hover | action / on-action / action-hover (also link, focus) |
| border / border-strong | border (decorative) / border-strong (control borders) |
| disabled-bg / disabled-text | disabled-bg / disabled-text |
| success, warning, info (-text/-bg) | same names; `error` -> `danger` |
| radius-control / card / dialog; space-N; font-body / display | same names |

`--nl-surface-subtle` is not consumed by any rule: light muted text on it is 4.45:1 (02 §2.3).

## App shell (opt-in)

A theme opts in with `--app-shell: 1`; only northlake does. Without it the header is `display:none`,
`data-scheme` is never set, and none of the header or icon JS runs.

- `<head>` script resolves System/Light/Dark (`localStorage` key `northlake-theme`, safe when storage
  is blocked) into `html[data-scheme]` before first paint and defines `window.appScheme_`.
- Header (`index.html` `<header class="app-brand">`): `a.app-home` (flame mark + app name, links to
  `./`), then `.app-tools`: theme control (inside the Menu panel when narrow), identity control
  (`#appAccount`, always its own header control, filled by `renderAccount_`; absent without
  Dispatch), Menu button. The narrow layout is a container query on the header (519px), not a device
  breakpoint. Signed out: a "Sign in" toggle opens a panel holding Google's button. Signed in:
  avatar or initials; name, email and Sign out only in the panel.
- Icons: Lucide outline paths inline in `SHELL_ICONS_`; `shellIconize_` swaps the emoji only when the
  shell is on, so PROD keeps its emoji. Nothing loads at runtime.
- Auth: static `index.html` on SIT and NUUC (optional Google sign-in via NUUC-Dispatch). PROD has none.
