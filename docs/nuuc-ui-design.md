# **Northlake Digital Design System**

Draft v0.1

Option A — Lake & Evergreen · November 2026

This specification translates Option A from the visual design review into practical standards for Northlake web applications. It covers colors, typography, layouts, components, icons, and accessibility.

**The existing Northlake logo is a fixed asset and must not be redesigned.** The UI should complement it without reproducing or reinterpreting its flame.

## **Color system**

Primary interactive blue

Primary actions, active navigation, links

\#1663A5

Brand lake blue

Accents and decorative highlights

\#1E88E5

Leaf green

Brand accents and supporting graphics

\#7CB342

Sunset orange

Highlights and warm accents

\#FB8C00

Golden yellow

Decorative highlights

\#FDD835

Charcoal

Headings and prominent text

\#212121

Canvas

Application background

\#F4F7F6

White

Cards, panels and content surfaces

\#FFFFFF

The brand colors originate from the existing draft. The darker interactive blue and neutral UI surfaces are Option A proposals, not yet approved standards.

## **Visual component examples**

**Action Items**

Northlake application

\+ New item

View all

Review volunteer schedule

Review

Update equipment inventory

In progress

Confirm room setup

Open

Illustrative UI treatment. Production components should use the semantic design tokens specified below.

# Northlake Digital Design System

**Implementation Specification — Draft 0.1**

**Visual direction:** Option A — Lake & Evergreen\
**Status:** Proposed standard, subject to Communications Committee review\
**Scope:** Northlake-developed web applications and related digital tools

## 1. Purpose

Establish a consistent, recognizable, accessible UI design language for applications developed or maintained on behalf of Northlake Unitarian Universalist Church.

The system should communicate Northlake's existing identity while remaining practical for forms, dashboards, administrative tools, volunteer applications, and public-facing interfaces.

**Design principles:**

- Welcoming, clear, calm, and inclusive.
- Consistent with Northlake's existing brand identity.
- Functionality and readability before decoration.
- Accessible and responsive.
- Reusable and easy to maintain across applications.

## 2. Brand Identity

Northlake's existing logos and flame artwork are authoritative.

**Requirements:**

- Use approved, unmodified Northlake logo assets.
- Do not redraw, recolor, simplify, or reinterpret the flame.
- Preserve original proportions and transparent backgrounds.
- Never stretch, crop, or overlay other graphic elements on logos.
- Use the original square logo for compact application branding where appropriate.
- Use an existing horizontal logo variant when the header requires a wide format.
- Keep sufficient clear space around the logo.
- Do not create new flame-themed icons to replace standard application icons.

**Authoritative asset location:**

[https://drive.google.com/drive/u/0/folders/1gF2pF8IIT3Y0mZUWtIY2tDRLUuOE-Z_A](https://drive.google.com/drive/u/0/folders/1gF2pF8IIT3Y0mZUWtIY2tDRLUuOE-Z_A)

Recommended initial asset: `NorthlakeLogoOriginalSquare.png`.

Obtain approved assets from the shared folder and include them in the application's static assets. Do not depend on an authenticated Google Drive URL for runtime rendering.

## 3. Design Tokens

Applications should use semantic CSS custom properties rather than hardcoding colors throughout their components.

### 3.1 Brand and UI colors

| Token                | Hex       | Purpose                                   |
| -------------------- | --------- | ----------------------------------------- |
| `--nl-primary`       | `#1663A5` | Primary actions, active navigation, links |
| `--nl-primary-hover` | `#104E81` | Primary hover state                       |
| `--nl-primary-soft`  | `#E9F3FB` | Selected items, subtle highlights         |
| `--nl-brand-blue`    | `#1E88E5` | Brand accent                              |
| `--nl-brand-green`   | `#7CB342` | Brand accent                              |
| `--nl-brand-orange`  | `#FB8C00` | Brand accent                              |
| `--nl-brand-yellow`  | `#FDD835` | Brand accent                              |
| `--nl-brand-red`     | `#E53935` | Brand accent                              |
| `--nl-text`          | `#24313B` | Primary text                              |
| `--nl-text-muted`    | `#63717A` | Secondary text                            |
| `--nl-heading`       | `#212121` | Headings                                  |
| `--nl-background`    | `#F4F7F6` | Page background                           |
| `--nl-surface`       | `#FFFFFF` | Cards, dialogs, forms                     |
| `--nl-border`        | `#DCE4E5` | Borders and separators                    |

The stronger blue (`#1663A5`) is the primary interactive color. The original brighter lake blue remains a brand accent rather than the default button background.

### 3.2 Semantic states

State colors must represent meaning, independently of brand colors.

| State                     | Visual treatment                  |
| ------------------------- | --------------------------------- |
| Success / Completed       | Green                             |
| Warning / Needs attention | Amber                             |
| Error / Destructive       | Red                               |
| Information / Review      | Blue                              |
| Neutral / Unknown         | Gray                              |
| In progress               | Amber or blue with explicit label |

Do not use the same status vocabulary inconsistently between applications. For example, *Open* should not be confused with *Completed* simply because both could be represented in green.

Use both text and iconography, not color alone, to convey state.

Final semantic colors must pass contrast testing before production use.

### 3.3 Typography

| Usage                 | Font                  | Suggested size |
| --------------------- | --------------------- | -------------- |
| Page title            | Georgia, serif        | 30–36px        |
| Section heading       | Georgia, serif        | 24–28px        |
| Card heading          | Open Sans, sans-serif | 18–20px        |
| Body text             | Open Sans, sans-serif | 16px           |
| Labels and buttons    | Open Sans, sans-serif | 14–16px        |
| Secondary information | Open Sans, sans-serif | 14px           |

Guidelines:

- Default to a sans-serif for all interactive controls and data-heavy screens.
- Use Georgia selectively for primary headings and editorial content.
- Use normal sentence case, not all capitals, for most headings and labels.
- Maintain readable line spacing of approximately 1.5 for body text.
- Use system font fallbacks if the preferred fonts are unavailable.
- Avoid decorative typography in administrative interfaces.

### 3.4 Spacing and geometry

Use a consistent 4px spacing scale:

`4, 8, 12, 16, 24, 32, 48, 64px`

Recommended geometry:

| Element                  | Radius        |
| ------------------------ | ------------- |
| Buttons and input fields | 8px           |
| Cards and panels         | 12px          |
| Dialogs                  | 16px          |
| Pills / status labels    | Fully rounded |

Use subtle borders and restrained shadows. Avoid heavy gradients, glossy effects, or excessive decoration.

Prefer white content surfaces over a light neutral page background.

## 4. Application Layout

### Header

- Display an approved Northlake logo and the application name.
- Provide clearly labeled navigation.
- Highlight the active section with primary blue.
- Keep branding consistent across applications.
- Collapse navigation appropriately on small screens.

### Main content

- Recommended maximum width: 1200px for ordinary application pages.
- Allow wider layouts for complex tables and dashboards.
- Use consistent horizontal page padding.
- Provide a clear page heading.
- Place primary actions in predictable locations.
- Group related content into sections or cards.

### Footer

Keep application footers minimal. Include appropriate organizational identification, support information, and accessibility or privacy links where relevant.

## 5. UI Components

### Buttons

**Primary:** Solid `--nl-primary` background with white text.

**Secondary:** White background, primary-color border, and primary-color text.

**Tertiary:** Text-only control with an appropriate hover and focus treatment.

**Destructive:** Explicit red styling and a confirmation step where the consequence warrants one.

Requirements:

- Minimum 44×44px interactive target where practical.
- Visible keyboard focus indicator.
- Hover, active, disabled, and loading states.
- Use action-oriented labels such as “Save changes,” not vague labels such as “OK.”

### Forms

- Place visible labels above or beside inputs.
- Do not use placeholders as the only field labels.
- Use consistent 8px input corner radii.
- Show required fields clearly.
- Display validation errors near the affected field.
- Preserve entered values after validation errors.
- Present successful saves or submissions with clear feedback.

### Cards

Cards should have:

- White background.
- Light neutral border.
- Approximately 12px border radius.
- Consistent internal spacing of 16–24px.
- Clear heading and content hierarchy.

Use cards to group related information, not as a default wrapper for every individual field.

### Tables and lists

- Prefer clear horizontal separators.
- Right-align numeric values.
- Maintain readable row spacing.
- Support sorting and filtering where useful.
- Handle narrow screens with appropriate reflow or horizontal scrolling.
- Never communicate important information through a colored cell alone.

### Alerts and status indicators

Use consistent styling for information, success, warning, and error messages.

Each alert should contain enough text to explain what happened and, where relevant, what the user should do.

## 6. Iconography

**Recommended library: Lucide**, using SVG icons.

Conventions:

- Default size: 20px.
- Compact controls: 16px.
- Prominent actions: 24px.
- Consistent outline style and stroke weight.
- Use the same icon for the same action across applications.
- Pair unfamiliar icons with visible labels.
- Provide accessible names for icon-only buttons.

Examples:

| Action   | Recommended icon |
| -------- | ---------------- |
| Home     | `house`          |
| Calendar | `calendar-days`  |
| People   | `users`          |
| Settings | `settings`       |
| Edit     | `pencil`         |
| Delete   | `trash-2`        |
| Search   | `search`         |
| Save     | `save`           |
| Review   | `eye`            |
| Complete | `check-circle`   |
| Help     | `circle-help`    |

The icon library is a technical recommendation, not an existing approved Northlake standard.

## 7. Imagery

Use real photographs of Northlake activities and the community when suitable and authorized.

Preferred subjects:

- Community gatherings and participation.
- Worship and reflection.
- Pacific Northwest landscapes and nature.
- Genuine volunteer and program activities.

Avoid generic corporate imagery that does not reflect Northlake.

Use consistent image treatment:

- Natural colors.
- Appropriate cropping.
- Rounded corners when displayed in cards.
- Meaningful alternative text.
- Optimized image sizes for web delivery.

Follow Northlake media privacy and copyright requirements, including permissions for identifiable minors.

Do not use decorative flame artwork invented by individual applications.

## 8. Accessibility and Responsiveness

Target WCAG 2.2 Level AA.

Requirements:

- Minimum 4.5:1 contrast for normal text.
- Minimum 3:1 contrast for large text and relevant non-text UI indicators.
- Fully keyboard-operable controls.
- Visible focus indicators.
- Accessible form labels and error descriptions.
- Respect reduced-motion preferences.
- Support text resizing up to 200% without loss of functionality.
- Ensure responsive layouts work at 320px CSS viewport width without unnecessary horizontal scrolling.
- Use semantic HTML and appropriate accessible names.
- Do not rely solely on hover for essential information.

Check the proposed colors in actual component contexts. Brand colors are not automatically valid as accessible text or status colors.

## 9. CSS Reference Implementation

Use these tokens as the starting point for a shared stylesheet or theme configuration.

```
:root {
  /* Core UI */
  --nl-primary: #1663a5;
  --nl-primary-hover: #104e81;
  --nl-primary-soft: #e9f3fb;

  /* Brand */
  --nl-brand-blue: #1e88e5;
  --nl-brand-green: #7cb342;
  --nl-brand-orange: #fb8c00;
  --nl-brand-yellow: #fdd835;
  --nl-brand-red: #e53935;

  /* Surfaces and text */
  --nl-text: #24313b;
  --nl-text-muted: #63717a;
  --nl-heading: #212121;
  --nl-background: #f4f7f6;
  --nl-surface: #ffffff;
  --nl-border: #dce4e5;

  /* Typography */
  --nl-font-body: "Open Sans", Arial, sans-serif;
  --nl-font-display: Georgia, serif;

  /* Geometry */
  --nl-radius-control: 8px;
  --nl-radius-card: 12px;
  --nl-radius-dialog: 16px;

  /* Spacing */
  --nl-space-1: 4px;
  --nl-space-2: 8px;
  --nl-space-3: 12px;
  --nl-space-4: 16px;
  --nl-space-6: 24px;
  --nl-space-8: 32px;
}

body {
  background: var(--nl-background);
  color: var(--nl-text);
  font: 16px/1.5 var(--nl-font-body);
}

h1, h2 {
  color: var(--nl-heading);
  font-family: var(--nl-font-display);
  font-weight: 400;
}

.nl-button-primary {
  background: var(--nl-primary);
  color: #fff;
  border: 1px solid var(--nl-primary);
  border-radius: var(--nl-radius-control);
  padding: 10px 16px;
  font: 600 1rem var(--nl-font-body);
  cursor: pointer;
}

.nl-button-primary:hover {
  background: var(--nl-primary-hover);
}

.nl-card {
  background: var(--nl-surface);
  border: 1px solid var(--nl-border);
  border-radius: var(--nl-radius-card);
  padding: var(--nl-space-6);
}

:where(button, a, input, select, textarea):focus-visible {
  outline: 3px solid var(--nl-primary);
  outline-offset: 3px;
}
```

Production teams should implement disabled, loading, hover, focus, validation, and semantic status variants as part of their component library.

## 10. Integration Guidelines

This specification is framework-independent. It may be used with React, Vue, plain HTML/CSS/JavaScript, or a WordPress theme.

Implementation recommendations:

1. Centralize tokens in one stylesheet or theme configuration.
2. Use reusable UI components where practical.
3. Avoid hardcoded styling scattered across application pages.
4. Keep original Northlake logos as version-controlled static assets.
5. Retain existing application behavior and workflows while updating styling.
6. Do not replace accessible native controls with custom controls unless there is a clear usability need.
7. Use responsive layouts from the beginning.
8. Verify that visual changes do not introduce regressions.

Do not introduce a new UI framework solely to apply this design language.

## 11. Acceptance Criteria

A Northlake application conforms to this draft when:

- It uses an approved, unmodified Northlake logo.
- Its primary UI styling follows Option A.
- Colors are implemented through centralized design tokens.
- Typography follows the recommended hierarchy.
- Buttons, forms, navigation, and cards use consistent components.
- Icons follow a consistent SVG icon library.
- Semantic states are used consistently.
- Core workflows meet WCAG 2.2 AA requirements.
- Layouts remain usable on mobile and desktop.
- Existing application functionality is preserved.
- Screens are reviewed for consistency with the reference design.

## 12. Status and Governance

This document is an **implementation proposal**, not an adopted Northlake branding policy.

The Communications Committee should review and approve the visual identity choices, including colors, typography, and approved logo variants.

The application development team should maintain the technical design system, accessibility requirements, and reusable UI components.

Any future changes to the Northlake brand identity should be made through the Communications review process rather than independently by application developers.

**Reference materials**

- Draft NUUC Design Language 0.0.1: [https://docs.google.com/document/d/1KLNcfLGZ3m9-C_lRddMeaHHOEg9Yi0gdT-fy9Tm7gLo/edit](https://docs.google.com/document/d/1KLNcfLGZ3m9-C_lRddMeaHHOEg9Yi0gdT-fy9Tm7gLo/edit)
- Northlake Logos and Images: [https://drive.google.com/drive/u/0/folders/1gF2pF8IIT3Y0mZUWtIY2tDRLUuOE-Z_A](https://drive.google.com/drive/u/0/folders/1gF2pF8IIT3Y0mZUWtIY2tDRLUuOE-Z_A)
- Northlake website: [https://northlakeuu.org/](https://northlakeuu.org/)

---

**End of Draft 0.1**

## **Recommended implementation approach**

For an existing web app, I would use this specification in two passes.

**First pass — Visual consistency:** Introduce design tokens, apply the original logo, standardize typography, buttons, cards, navigation, and icons. Preserve the existing user workflows.

**Second pass — Accessibility and refinement:** Audit contrast, keyboard behavior, responsive layouts, status semantics, and inconsistent controls. Correct anything that fails accessibility or creates usability problems.

This avoids unnecessarily rebuilding applications just to standardize their appearance.

Because this is still a draft, I would ask Communications to approve the palette and typography before directing developers to treat those choices as mandatory.


