# Northlake Digital Design System

**Web Application Design Specification — Draft 0.3**

**Visual direction:** Option A — Lake & Evergreen\
**Status:** Proposed standard, pending Communications Committee review\
**Updated:** November 21, 2026\
**Applies to:** Northlake-developed web applications and related digital interfaces

---

## 1. Purpose and Design Principles

Establish a consistent, recognizable, accessible design language for web applications developed or maintained on behalf of Northlake Unitarian Universalist Church.

Applications should feel like part of the same Northlake family without unnecessarily restricting their functionality or requiring a particular technology stack.

### Design principles

1. **Recognizably Northlake.** Preserve the church's established logo, colors, and visual identity.
2. **Welcoming and approachable.** Use clear language, open layouts, and restrained styling.
3. **Clarity before decoration.** Prioritize usability and readability.
4. **Consistent.** Use common conventions across applications.
5. **Accessible.** Support users with varied visual, motor, and cognitive abilities.
6. **Adaptable.** Support light and dark themes and responsive layouts.
7. **Maintainable.** Centralize visual definitions and reuse components.

### Scope

This specification defines how common UI elements should look and behave **when present**. It does not require applications to implement functionality they do not need.

Application-specific requirements determine features, authentication, workflows, data access, and authorization.

---

## 2. Brand Identity and Logo Usage

### 2.1 Authoritative logo assets

Northlake's existing logo is a fixed organizational identity asset.

**Authoritative source:**

[Northlake Logos and Images — Communications Drive](https://drive.google.com/drive/u/0/folders/1gF2pF8IIT3Y0mZUWtIY2tDRLUuOE-Z_A)

Preferred asset for application branding:

`Northlake-logo-1024x3031 transparent.png`

Other existing assets include:

- `NorthlakeLogoOriginalSquare.png`
- `Northlake-white-on-transparent.png`
- Existing horizontal, square, and banner variants.

### 2.2 Logo requirements

- Use approved, unmodified Northlake artwork.
- Prefer transparent-background logos for application headers.
- Do not redraw, recolor, simplify, or reinterpret the flame.
- Preserve aspect ratio and appropriate clear space.
- Never stretch, crop, distort, or apply decorative effects.
- Do not overlay text or other icons.
- Avoid backgrounds that compromise logo readability.

### 2.3 Theme-specific logo usage

**Light theme:** Use an approved transparent-background logo on a light surface.

**Dark theme:** Use an existing approved transparent variant that remains legible against the selected dark background.

The existing white-on-transparent variant is a candidate, subject to visual verification.

If no existing variant is readable, provide a deliberate contrasting logo surface instead of modifying the artwork.

### 2.4 Logo behavior

The logo should normally identify Northlake and, when clickable, navigate to the application's home page.

**Do not use the logo itself as a menu trigger.**

Navigation and identity controls should have their own clearly identifiable controls.

### 2.5 Deployment

Include approved logo files as application static assets rather than loading them directly from authenticated Google Drive URLs.

```
<img
  src="/assets/northlake-logo-transparent.png"
  alt="Northlake Unitarian Universalist Church"
  class="northlake-logo"
/>
```

```
.northlake-logo {
  display: block;
  height: 56px;
  width: auto;
  max-width: 100%;
  object-fit: contain;
}
```

Applications may adjust logo height to suit the header while preserving proportions.

---

## 3. Color System

### 3.1 Brand colors

| Brand color   | Hex       |
| ------------- | --------- |
| Flame Red     | `#E53935` |
| Sunset Orange | `#FB8C00` |
| Golden Yellow | `#FDD835` |
| Leaf Green    | `#7CB342` |
| Lake Blue     | `#1E88E5` |
| Charcoal      | `#212121` |

These are drawn from the existing draft Northlake design language.

Brand colors are not automatically suitable for functional UI elements. Interactive and semantic colors must meet accessibility requirements.

### 3.2 Semantic UI colors

| Semantic token          | Light theme | Dark theme |
| ----------------------- | ----------- | ---------- |
| Primary action          | `#1663A5`   | `#72B7EF`  |
| Primary hover           | `#104E81`   | `#9ACCF4`  |
| Text on primary         | `#FFFFFF`   | `#10232F`  |
| Primary soft background | `#E9F3FB`   | `#183B53`  |
| Primary soft text       | `#1663A5`   | `#B9DFFF`  |
| Page background         | `#F4F7F6`   | `#17232B`  |
| Card / panel            | `#FFFFFF`   | `#22323D`  |
| Subtle surface          | `#EDF2F2`   | `#2B3D49`  |
| Primary text            | `#24313B`   | `#EAF1F5`  |
| Secondary text          | `#63717A`   | `#B5C6D0`  |
| Headings                | `#212121`   | `#F5F8FA`  |
| Borders                 | `#DCE4E5`   | `#405361`  |

These are proposed UI values. Validate contrast in actual components before finalizing them.

### 3.3 Usage

- Use primary blue for major interactive elements.
- Use green, orange, and yellow selectively as supporting accents.
- Prefer neutral backgrounds for forms and data-heavy interfaces.
- Reserve semantic red for errors and destructive actions.
- Do not communicate state through color alone.
- Use semantic tokens rather than hardcoded colors.

### 3.4 Status meanings

| Meaning                   | Color family  |
| ------------------------- | ------------- |
| Success / Completed       | Green         |
| Warning / Needs attention | Amber         |
| Error / Failed            | Red           |
| Information / Review      | Blue          |
| Neutral / Unknown         | Gray          |
| In progress               | Blue or amber |

Status indicators must include descriptive text and, where useful, icons.

Application-specific workflow states should be mapped consistently. For example, *Open* does not necessarily mean *Success*.

---

## 4. Light and Dark Theme Support

### 4.1 Theme selection

Applications should support:

- **System** — Follow browser or operating-system preference.
- **Light** — Always use the light theme.
- **Dark** — Always use the dark theme.

System is the recommended default.

An explicit user preference overrides the system setting.

Where practical, persist that choice using application settings or browser storage.

### 4.2 Theme behavior

- Apply theme changes without reloading where practical.
- Do not reset form inputs or application state.
- Theme all components, including dialogs, charts, tables, and notifications.
- Avoid a visible flash of the incorrect theme on startup.
- Respect reduced-motion preferences.
- Set the appropriate CSS `color-scheme`.

### 4.3 Theme-dependent styling

Use shared semantic tokens for both themes.

Do not assume that:

- Backgrounds are white.
- Text is always dark.
- Primary buttons always use white text.
- Borders use the same color in both themes.

### 4.4 Logo behavior

Applications may select different approved logo variants for light and dark themes.

The logo's artwork must never be recolored or modified programmatically.

---

## 5. Typography

### Font families

**Primary interface font:** Open Sans, with Arial and system sans-serif fallbacks.

**Display font:** Georgia or an approved equivalent serif.

Use sans-serif for navigation, controls, tables, forms, and most application content.

Use serif typography selectively for prominent page headings and editorial content.

### Type scale

| Element            | Suggested size |
| ------------------ | -------------- |
| Page title         | 30–36px        |
| Section heading    | 24–28px        |
| Card heading       | 18–20px        |
| Body               | 16px           |
| Buttons and labels | 14–16px        |
| Secondary text     | 14px           |

Use approximately 1.5 line height for body text.

Prefer sentence case over all-capital labels.

---

## 6. Spacing and Geometry

Use a consistent 4px-based spacing scale:

`4, 8, 12, 16, 24, 32, 48, 64px`

### Corner radii

| Component          | Radius        |
| ------------------ | ------------- |
| Buttons and inputs | 8px           |
| Cards              | 12px          |
| Dialogs            | 16px          |
| Pills and badges   | Fully rounded |

Use subtle borders and restrained shadows.

Avoid heavy gradients, glossy effects, or unnecessary decorative elements.

---

## 7. Application Layout and Navigation

### 7.1 General page structure

Applications should generally provide:

- A recognizable application header.
- Clear application branding.
- Predictable navigation.
- A clearly identified main content region.
- Consistent placement of primary actions.
- Related content grouped logically.
- A minimal footer where appropriate.

Typical content width should be approximately 1200px or less, except where wider layouts benefit complex dashboards or tables.

### 7.2 Header layout — Wide screens

Preferred arrangement:

**Left:** Approved Northlake logo and application name.

**Middle or left navigation area:** Primary application navigation, where appropriate.

**Right:** Optional help, identity, and account controls.

Avoid displaying controls that the application does not support.

### 7.3 Header layout — Narrow screens

When the available width is insufficient for separate navigation and identity controls:

1. Preserve a recognizable Northlake logo and application identity.
2. Collapse navigation into a clearly labeled menu control, typically using a three-bar icon.
3. Prefer placing the menu button in the upper-right corner, opposite the logo.
4. Applications with established left-side navigation may use an upper-left menu.
5. Combine application navigation and identity controls within the menu where useful.
6. Keep identity actions visually separated from ordinary navigation links.
7. Collapse controls according to available space rather than a mandatory device-width breakpoint.

Do not allow header controls to wrap awkwardly, overlap, or become clipped.

### 7.4 Mobile menu organization

When account controls share a mobile navigation menu, use this general order:

1. Application navigation.
2. Optional help or supporting actions.
3. Account identity and personal preferences, when supported.
4. Sign in or Sign out, when applicable.

A divider or similar visual separation should distinguish account controls from application navigation.

### 7.5 Menu accessibility

Responsive menus must:

- Have an accessible label, such as “Open navigation menu.”
- Indicate their expanded or collapsed state.
- Support keyboard and touch operation.
- Support Escape to close where applicable.
- Manage keyboard focus appropriately.
- Avoid trapping focus incorrectly.
- Keep essential functionality discoverable.

Do not attach hidden navigation to the Northlake logo.

---

## 8. User Identity and Account Navigation

### 8.1 Purpose

Establish consistent placement and presentation of identity and account controls **when an application supports them**.

This section does not prescribe whether users must authenticate.

### 8.2 Authentication models

Northlake applications may:

- Operate entirely without user authentication.
- Provide public functionality with optional sign-in.
- Require authentication for selected features.
- Require authentication for all functionality.

The application determines its authentication model and access requirements.

The design system does not impose one.

### 8.3 Signed-out presentation

When authentication is available, provide a clearly labeled **Sign in** control.

On wide screens, prefer the upper-right header.

On narrow screens, Sign in may appear within the collapsed navigation menu.

Applications should not imply that sign-in is required when anonymous access is supported.

### 8.4 Signed-in presentation

When authenticated, display an identifiable account control.

It may contain:

- Display name.
- Initials.
- Avatar.
- Appropriate account icon.

Selecting the control should reveal supported account actions.

Typical actions include:

- Account or profile information.
- Personal preferences.
- Sign out.

Only show capabilities actually supported by the application.

### 8.5 Account menu conventions

- Use **Sign out** as the explicit label for ending the application session.
- Do not rely on an unlabeled exit icon.
- Display the actual authenticated identity where identity information is shown.
- Do not assume users have `northlakeuu.org` email addresses.
- Keep personal preferences separate from application administration.
- Only offer role switching if supported by the application.

### 8.6 Application-specific decisions

The design specification does not determine:

- Whether authentication is required.
- Which features anonymous users may access.
- When sign-in is requested.
- Which identity providers are supported.
- Whether accounts have profiles or roles.
- How authentication affects application state.
- Session duration or authorization rules.

These belong to the application's functional and security requirements.

### 8.7 Responsive identity behavior

When account information cannot fit comfortably within the header:

- Move it into the responsive navigation menu.
- Preserve a clear Sign in or Sign out label.
- Group account information separately from application functions.
- Avoid shrinking text or compressing controls excessively.

On desktop, the account control should generally remain in the upper-right header.

### 8.8 Accessibility

- Account menu controls require accessible names.
- Keyboard users must be able to open, navigate, and close menus.
- Menu focus behavior must be predictable.
- Icons must not be the sole indication of account state.
- Identity controls must remain usable in both themes.

---

## 9. Standard UI Components

### 9.1 Buttons

**Primary:** Filled using the theme's primary action color.

**Secondary:** Neutral background with primary-colored text and border.

**Tertiary:** Minimal text-oriented styling.

**Destructive:** Explicit destructive styling, with confirmation when appropriate.

Support default, hover, focus, active, disabled, and loading states.

### 9.2 Forms

- Provide visible labels.
- Clearly identify required fields.
- Show errors near the relevant input.
- Preserve entered data after validation failures.
- Use consistent spacing and geometry.
- Provide clear submission feedback.
- Support keyboard navigation and assistive technology.

### 9.3 Cards

Use:

- Theme-aware surface color.
- Subtle border.
- 12px corner radius.
- Consistent 16–24px padding.
- Clear content hierarchy.

Avoid unnecessary nesting.

### 9.4 Tables

- Maintain readable row spacing.
- Right-align numeric data.
- Use subtle separators.
- Support responsive layouts.
- Provide sorting and filtering where useful.
- Show meaningful empty states.
- Use labels as well as colors for status.

### 9.5 Notifications and dialogs

- Use semantic state styling.
- Provide actionable messages.
- Support keyboard interaction.
- Manage focus correctly.
- Confirm destructive actions when appropriate.

### 9.6 Optional components

Search, notifications, account management, theme controls, and other common features should use consistent styling when implemented.

The design system does **not** require every application to provide these features.

---

## 10. Iconography

### Recommended icon library

**Lucide SVG icons**

Use a consistent outline style and stroke weight.

### Size conventions

| Usage     | Size |
| --------- | ---- |
| Compact   | 16px |
| Standard  | 20px |
| Prominent | 24px |

### Suggested mappings

| Concept         | Icon                |
| --------------- | ------------------- |
| Home            | `house`             |
| Calendar        | `calendar-days`     |
| People          | `users`             |
| Settings        | `settings`          |
| Edit            | `pencil`            |
| Delete          | `trash-2`           |
| Search          | `search`            |
| Save            | `save`              |
| Review          | `eye`               |
| Completed       | `circle-check`      |
| Help            | `circle-help`       |
| Warning         | `triangle-alert`    |
| Sign in         | `log-in`            |
| Sign out        | `log-out`           |
| Account         | `circle-user-round` |
| Navigation menu | `menu`              |
| Light theme     | `sun`               |
| Dark theme      | `moon`              |
| System theme    | `monitor`           |

Use visible text labels for unfamiliar icons.

Do not introduce invented Northlake flame icons for ordinary application functions.

---

## 11. Imagery

Preferred imagery should reflect:

- Northlake gatherings and community.
- Worship and reflection.
- Volunteer activities.
- Pacific Northwest landscapes.
- Nature and natural light.

Guidelines:

- Prefer authentic imagery.
- Avoid unnecessary filters.
- Preserve natural image proportions.
- Use consistent cropping and placement.
- Provide meaningful alternative text.
- Optimize image performance.
- Respect copyright and privacy requirements.

Follow Northlake's media privacy policies, particularly regarding identifiable minors.

---

## 12. Accessibility

Applications should target **WCAG 2.2 Level AA**.

Requirements include:

- At least 4.5:1 contrast for normal text.
- At least 3:1 contrast for large text and applicable non-text indicators.
- Keyboard-operable controls.
- Visible focus indicators.
- Accessible names and labels.
- Meaningful semantic HTML.
- Support for 200% text resizing.
- Responsive reflow.
- Appropriate pointer target sizes.
- Reduced-motion support.
- No essential information communicated only through color.

Both light and dark themes must be tested independently.

Navigation and identity menus must also meet these requirements.

---

## 13. CSS Design Token Reference

```
/* Northlake Digital Design System — Draft 0.3 */

:root {
  color-scheme: light;

  /* Brand colors */
  --nl-brand-red: #e53935;
  --nl-brand-orange: #fb8c00;
  --nl-brand-yellow: #fdd835;
  --nl-brand-green: #7cb342;
  --nl-brand-blue: #1e88e5;
  --nl-brand-charcoal: #212121;

  /* Light theme */
  --nl-primary: #1663a5;
  --nl-primary-hover: #104e81;
  --nl-on-primary: #ffffff;
  --nl-primary-soft: #e9f3fb;
  --nl-primary-soft-text: #1663a5;

  --nl-background: #f4f7f6;
  --nl-surface: #ffffff;
  --nl-surface-subtle: #edf2f2;

  --nl-text: #24313b;
  --nl-text-muted: #63717a;
  --nl-heading: #212121;
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
  --nl-space-12: 48px;
  --nl-space-16: 64px;
}

:root[data-theme="dark"] {
  color-scheme: dark;

  --nl-primary: #72b7ef;
  --nl-primary-hover: #9accf4;
  --nl-on-primary: #10232f;
  --nl-primary-soft: #183b53;
  --nl-primary-soft-text: #b9dfff;

  --nl-background: #17232b;
  --nl-surface: #22323d;
  --nl-surface-subtle: #2b3d49;

  --nl-text: #eaf1f5;
  --nl-text-muted: #b5c6d0;
  --nl-heading: #f5f8fa;
  --nl-border: #405361;
}

body {
  margin: 0;
  background: var(--nl-background);
  color: var(--nl-text);
  font: 16px/1.5 var(--nl-font-body);
}

h1, h2 {
  color: var(--nl-heading);
  font-family: var(--nl-font-display);
  font-weight: 400;
}

h3, h4, h5, h6 {
  color: var(--nl-heading);
  font-family: var(--nl-font-body);
}

a {
  color: var(--nl-primary);
}

.nl-logo {
  display: block;
  height: 56px;
  width: auto;
  max-width: 100%;
  object-fit: contain;
}

.nl-button-primary {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  min-height: 44px;
  padding: 10px 16px;

  background: var(--nl-primary);
  color: var(--nl-on-primary);
  border: 1px solid var(--nl-primary);
  border-radius: var(--nl-radius-control);

  font: 600 1rem var(--nl-font-body);
  cursor: pointer;
}

.nl-button-primary:hover:not(:disabled) {
  background: var(--nl-primary-hover);
  border-color: var(--nl-primary-hover);
}

.nl-button-primary:disabled {
  opacity: 0.55;
  cursor: not-allowed;
}

.nl-button-secondary {
  min-height: 44px;
  padding: 10px 16px;

  background: var(--nl-surface);
  color: var(--nl-primary);
  border: 1px solid var(--nl-primary);
  border-radius: var(--nl-radius-control);

  font: 600 1rem var(--nl-font-body);
  cursor: pointer;
}

.nl-card {
  background: var(--nl-surface);
  color: var(--nl-text);
  border: 1px solid var(--nl-border);
  border-radius: var(--nl-radius-card);
  padding: var(--nl-space-6);
}

.nl-input {
  width: 100%;
  min-height: 44px;
  box-sizing: border-box;
  padding: 10px 12px;

  background: var(--nl-surface);
  color: var(--nl-text);
  border: 1px solid var(--nl-border);
  border-radius: var(--nl-radius-control);

  font: inherit;
}

:where(
  button,
  a,
  input,
  select,
  textarea
):focus-visible {
  outline: 3px solid var(--nl-primary);
  outline-offset: 3px;
}

@media (prefers-reduced-motion: reduce) {
  *,
  *::before,
  *::after {
    scroll-behavior: auto !important;
    animation-duration: 0.01ms !important;
    transition-duration: 0.01ms !important;
  }
}
```

The stylesheet establishes the base tokens and common styles. It is not a complete component library.

Applications must implement and validate additional component states, semantic status treatments, and responsive menu behavior.

---

## 14. Theme Selection Reference

Applications may implement theme selection using their existing state-management mechanism.

```
const STORAGE_KEY = "northlake-theme";

const media = window.matchMedia(
  "(prefers-color-scheme: dark)"
);

function getPreference() {
  try {
    const value = localStorage.getItem(STORAGE_KEY);
    if (["system", "light", "dark"].includes(value)) {
      return value;
    }
  } catch {
    // Storage may be unavailable.
  }

  return "system";
}

function applyTheme(preference) {
  const theme = preference === "system"
    ? (media.matches ? "dark" : "light")
    : preference;

  document.documentElement.dataset.theme = theme;
}

function setTheme(preference) {
  if (!["system", "light", "dark"].includes(preference)) {
    return;
  }

  try {
    localStorage.setItem(STORAGE_KEY, preference);
  } catch {
    // Continue without persistence.
  }

  applyTheme(preference);
}

applyTheme(getPreference());

media.addEventListener("change", () => {
  if (getPreference() === "system") {
    applyTheme("system");
  }
});
```

Applications should apply the stored theme as early as practical to avoid a flash of the incorrect theme.

---

## 15. Application Integration

### Existing applications

Adopt the design system incrementally.

**Phase 1 — Foundation**

- Install approved transparent logo assets.
- Introduce semantic color tokens.
- Add theme support.
- Standardize typography and spacing.

**Phase 2 — Shared visual conventions**

- Update headers and navigation.
- Standardize responsive menus.
- Apply consistent identity-control placement when applicable.
- Update buttons, forms, cards, and dialogs.
- Standardize iconography.

**Phase 3 — Validation**

- Review both themes.
- Verify desktop and mobile layouts.
- Test keyboard accessibility.
- Check contrast and other accessibility requirements.
- Verify existing functionality remains intact.

### New applications

Use this specification from initial development.

Reuse existing Northlake components where practical instead of independently redefining their appearance or behavior.

### Exceptions

Applications may depart from conventions when justified by their requirements.

Examples include:

- Specialized AV control interfaces.
- Complex data dashboards.
- Highly constrained displays.
- Workflows with unusual navigation requirements.

Document significant exceptions.

Do not compromise accessibility or modify the established Northlake logo.

---

## 16. Acceptance Criteria

### Brand identity

- Uses approved, unmodified Northlake logo assets.
- Uses transparent-background artwork where appropriate.
- Logo is readable in both themes.
- Logo is not used as a hidden menu trigger.

### Colors and typography

- Uses centralized semantic design tokens.
- Uses consistent typography and spacing.
- Supports light and dark styling.
- Avoids hardcoded colors that break theme support.

### Theme support

- Light theme functions correctly.
- Dark theme functions correctly.
- System preference is supported.
- Explicit preferences are preserved when practical.
- All applicable components respond consistently.

### Navigation

- Application branding appears consistently.
- Navigation is discoverable.
- Narrow headers collapse gracefully.
- Menu controls are clearly labeled.
- Responsive menus support keyboard and touch interaction.

### Identity controls, where applicable

- Sign-in and account controls use consistent placement.
- Mobile account controls remain accessible when collapsed.
- Signed-in identity is accurately represented.
- Sign out is clearly labeled.
- Anonymous functionality remains available where the application permits it.
- No unsupported account features are displayed.

### Accessibility

- WCAG 2.2 AA requirements are met.
- Contrast is verified in both themes.
- Keyboard navigation works.
- Focus indicators are visible.
- Mobile layouts remain usable.
- Status indicators do not rely only on color.

### Functional integrity

- Existing workflows remain intact.
- Styling changes do not alter authorization behavior.
- Existing application state is preserved appropriately.
- No unnecessary UI framework migration is required.

---

## 17. Governance and Maintenance

This document is an implementation proposal, not an adopted Northlake branding policy.

### Communications Committee

Responsible for reviewing and approving:

- Brand identity and logo usage.
- Visual direction.
- Brand colors and typography.
- Major changes to visual identity.

### Application development teams

Responsible for:

- Implementing approved visual standards.
- Maintaining reusable UI components.
- Ensuring accessibility.
- Applying responsive behavior.
- Maintaining theme compatibility.
- Documenting application-specific exceptions.

### Change control

Updates to the design system should be versioned.

Brand identity changes require appropriate Communications review.

Technical implementation improvements that preserve the approved visual language may be managed by application development teams.

---

## 18. References

- [Draft NUUC Design Language 0.0.1](https://docs.google.com/document/d/1KLNcfLGZ3m9-C_lRddMeaHHOEg9Yi0gdT-fy9Tm7gLo/edit)
- [Northlake Logos and Images](https://drive.google.com/drive/u/0/folders/1gF2pF8IIT3Y0mZUWtIY2tDRLUuOE-Z_A)
- [Northlake website](https://northlakeuu.org/)

---

## Revision History

| Version | Changes                                                                                                                                                                                                                             |
| ------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0.1     | Initial Option A — Lake & Evergreen specification                                                                                                                                                                                   |
| 0.2     | Added transparent logo preference and light/dark theme support                                                                                                                                                                      |
| 0.3     | Added responsive navigation and identity conventions, mobile menu behavior, and explicit support for anonymous application use; clarified that the design system standardizes presentation without requiring optional functionality |

**End of Draft 0.3**

The principal change in v0.3 is the separation of **UI conventions from functional requirements**. Developers now have clear guidance for desktop and mobile headers, account menus, and anonymous access without being forced into a particular authentication or navigation model.


