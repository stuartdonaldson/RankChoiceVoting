To create a comprehensive design language tailored specifically for **Northlake Unitarian Universalist Church** (located in Kirkland, WA), we must design an aesthetic system that reflects their mission, theology, and local presence.

Unitarian Universalism (UU) is rooted in inclusion, environmental stewardship, spiritual inquiry, and community action. Northlake UU's branding builds directly upon these principles with a warm, Pacific Northwest sensibility.

---

## 1. Core Vision & Design Principles

* **Inclusive & Open (Radical Hospitality):** Clean, airy, readable interfaces that feel welcoming to diverse members, visitors, and seekers of all ages and backgrounds.
* **Warmth & Natural Harmony:** Drawing design cues from the Pacific Northwest, organic shapes, and community gathering spaces rather than cold tech minimalism.
* **Serene Clarity & Purpose:** Prioritizing effortless navigation for Sunday service schedules, community action groups, pastoral care, and events without clutter or visual noise.

---

## 2. Color Palette

The color system is derived directly from the classic **UU Flaming Chalice** identity combined with the natural landscape of Kirkland and Lake Washington.

```
       Primary Accent            Warm Gold / Ember           Evergreen Slate          Lake Water Blue
        [ #007791 ]                 [ #D97706 ]                [ #2A4D38 ]              [ #4B7B94 ]

```

### Brand & Interface Colors

* **Deep Chalice Teal (`#007791`):** Primary brand color for main headers, interactive navigation, and key buttons. Symbolizes Lake Washington, clarity, and peace.
* **Warm Gold / Ember (`#D97706`):** Primary accent color representing the **flame of the UU chalice**, light, hope, and community warmth. Used for call-to-action buttons (e.g., *Visit This Sunday*, *Donate*), badges, and highlights.
* **Evergreen Slate (`#2A4D38`):** Deep forest green for grounded structural elements, footer surfaces, and environmental/social justice badges.
* **Lake Water Blue (`#4B7B94`):** Soft muted accent for secondary links, event tags, and card headers.
* **Parchment White (`#FBF9F5`):** Off-white light background color. Warm and soft on the eyes compared to stark `#FFFFFF`.
* **Charcoal Text (`#2D3748`):** Deep gray for body copy to maintain comfortable readability and softness.

---

## 3. Typography

The typography pairs a **friendly, modern serif** for headlines (invoking spiritual reflection and classic community storytelling) with a **highly legible, warm sans-serif** for body text and navigation.

| Hierarchy Level | Font Family | Weight / Style | Primary Use Case |
| --- | --- | --- | --- |
| **Display / Headlines** | **Lora** or **Merriweather** | Bold / Regular (Serif) | Page titles, sermon titles, inspirational quotes, chalice lightings. |
| **Subheaders & UI** | **Plus Jakarta Sans** or **Inter** | Semi-Bold / Bold (Sans-serif) | Section titles, event cards, menu items, button labels. |
| **Body Copy** | **Plus Jakarta Sans** or **Open Sans** | Regular (16px–18px) | Order of service, blog posts, community updates, pastoral notes. |
| **Metadata & Details** | **Plus Jakarta Sans** | Medium / Monospace | Timestamps (e.g., `Sunday @ 10:30 AM`), venue details, contact info. |

---

## 4. Iconography & Visual Motif

Icons should be **soft, rounded, and welcoming** (2px line weight, rounded joins) rather than sharp or rigid.

### Iconography Style

* **The Flaming Chalice:** The core visual symbol of UU faith. Used thoughtfully in page headers, service cards, and spiritual resources.
* **Sunday Worship:** A stylized chalice, open book, or acoustic music note.
* **Social Justice / Community Action:** Hands holding a heart, a growing leaf, or interconnected community rings.
* **Location / Gathering:** A cozy pine tree motif, compass, or warm building pin.
* **Inclusion / Welcome:** Open doors, rainbow accents, or welcoming hands.

---

## 5. UI Layout & Component Guidelines

### A. The Navigation Header

* **Top Bar:** Off-white Parchment (`#FBF9F5`) background with the Northlake UU logo and Flaming Chalice on the left.
* **Action CTAs:** Prominent **Warm Gold (`#D97706`)** button: `[ Visit This Sunday ]` or `[ Join Live Stream ]`.

### B. Sunday Service / Event Cards

* **Visual Style:** Clean, rounded cards (`border-radius: 12px`) with soft drop shadows (`0 4px 12px rgba(0,0,0,0.05)`).
* **Color Bar:** A 4px vertical accent bar on the left of each card indicating the category:
* **Warm Gold**: Sunday Worship & Sermons
* **Evergreen**: Social Justice & Environmental Action
* **Lake Blue**: Small Groups & Fellowship


* **Typography:** Serif title in Deep Chalice Teal, time/date in bold sans-serif.

### C. Announcement & Quote Blocks

* **Blockquote Style:** Gentle background container (`#F0F4F6`), large serif quotation marks in Warm Gold, with a subtle border. Perfect for monthly ministry themes (e.g., *Pluralism*, *Interconnection*, *Equity*).

---

## 6. Design Tokens Reference (JSON Format)

```json
{
  "color": {
    "background-primary": "#FBF9F5",
    "surface-card": "#FFFFFF",
    "text-primary": "#2D3748",
    "text-muted": "#64748B",
    "brand-teal": "#007791",
    "brand-gold": "#D97706",
    "brand-evergreen": "#2A4D38",
    "brand-lake-blue": "#4B7B94"
  },
  "typography": {
    "font-display": "Lora, Merriweather, serif",
    "font-body": "Plus Jakarta Sans, sans-serif",
    "font-size-base": "16px"
  },
  "border-radius": {
    "button": "8px",
    "card": "12px",
    "pill": "9999px"
  },
  "shadows": {
    "card-soft": "0 4px 12px rgba(0, 0, 0, 0.05)"
  }
}

```

---

## Summary Comparison

Compared to F3's dark, high-contrast, tactical grit (Iron Charcoal `#1B1E21` and Tactical Red `#C83737`), the Northlake UU design language relies on **gentle light surfaces (Parchment `#FBF9F5`), Deep Teal (`#007791`), Warm Gold (`#D97706`), and inviting serif typography (Lora)**—creating an online environment that feels like stepping into a warm, inclusive community hall.
