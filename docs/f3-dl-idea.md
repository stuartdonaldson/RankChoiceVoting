To build an app that truly resonates with the men of F3 (the "PAX"), the design language needs to balance **tactical grit with modern digital usability**. F3’s roots are grounded in early-morning outdoor workouts, military-style discipline, local community, and authentic brotherhood.

The app shouldn't feel like a flashy consumer fitness tracker (like Strava or Nike Training Club) or a generic corporate tool (like basic Slack or Teams). Instead, it should feel **rugged, reliable, utility-driven, and distinctly "F3."**

Here is a proposed visual and UI Design Language for an F3 Mobile & Web Application.

---

## 1. Vision & Core Design Principles

* **Tactical Utility (Function First):** A man checking the app at 5:15 AM in a freezing, wet parking lot needs high contrast, massive hit targets, and zero fluff.
* **Brotherhood & Grit:** Rugged aesthetic—industrial, outdoorsy, grounded—without looking dated or cluttered.
* **The 3 Fs (Fitness, Fellowship, Faith):** Clear visual separation for the three pillars so men can easily pivot between workouts (1st F), social events/CSAUPs (2nd F), and community service or leadership (3rd F).

---

## 2. Color Palette

The color system combines F3 Nation’s iconic dark industrial roots with the moody, forested tones of the Pacific Northwest (Puget Sound / Cascades).

### Primary Colors

* **Iron Charcoal (`#1E2225`):** Main dark background. Dark mode by default for early-morning pre-blast checking.
* **Asphalt Black (`#121416`):** Secondary background, cards, and container surfaces.
* **Standard White (`#FFFFFF`):** High-contrast text and primary iconography.

### Secondary & Accent Colors

* **Gloom Gray (`#8A949E`):** Secondary text, subtle borders, inactive tab icons.
* **Tac-Orange / Blaze (`#FF5A36`):** The primary **Action Color**. Used for critical buttons (e.g., *HC - Hard Commit*, *Take the Q*), alerts, and urgent SLT announcements.
* **Sound Steel Blue (`#3B82F6`):** Accent color for **1st F (Fitness)** and AO workout map pins.
* **Fellowship Amber (`#F59E0B`):** Accent color for **2nd F (Fellowship)** gatherings, Happy Hours, and Co-working posts.
* **Faith Green (`#10B981`):** Accent color for **3rd F (Faith/Service)** projects, Q-Source discussions, and community initiatives.

---

## 3. Typography

F3 visual identity often relies on stencil or military-inspired typography for headlines paired with clean, readable sans-serifs for operational details (Backblasts, HC lists, rosters).

### Typeface Hierarchy

| Level | Font Family | Style / Weight | Primary Use Case |
| --- | --- | --- | --- |
| **Display / Headers** | **Black Ops One** or **Chakra Petch** | Bold / All Caps | Screen titles, AO names, "Gloom" headers, Badges. |
| **Subheaders / Titles** | **Inter** or **DM Sans** | Bold / Semi-Bold (18px–24px) | Card titles, Leader names (PAX), Role titles (Nantan, Site Q). |
| **Body & Logs** | **Inter** or **SF Pro / Roboto** | Regular / Medium (14px–16px) | Backblasts, threads, event descriptions, Slack-digest feeds. |
| **Tactical / Data** | **JetBrains Mono** or **Roboto Mono** | Bold (12px–14px) | Timestamps (e.g., `0530 PST`), AO Grid coordinates, Attendance counts, Permalinks. |

---

## 4. Iconography & Custom Symbols

Icons should be **bold, monoline or solid tactical shapes** (2px stroke minimum) so they remain legible on mobile screens outdoors in cold or direct sunlight.

### Core F3 Iconography Set

* **AO / Workout Location:** A compass or tactical map pin with an embedded shovel flag icon.
* **1st F (Fitness):** A stylized kettlebell or barbells.
* **2nd F (Fellowship):** Two clinking coffee mugs or a campfire icon.
* **3rd F (Faith/Service):** A hand reaching up or an anchor icon.
* **Take the Q:** A megaphone, whistle, or shield badge icon.
* **Hard Commit (HC):** A checkmark inside a heavy iron ring or a raised fist.
* **COT (Circle of Trust):** Concentric circles or a ring of stylized figures.
* **FNG (Friendly New Guy):** A "sprout" or "0.0" badge tag.

---

## 5. Layout & UI Components

### A. The "Gloom Bar" (Header)

* **Top Bar:** Displays the local sub-region (e.g., `F3 KIRKLAND` or `F3 CASCADES`), current temperature, weather icon, and the user's F3 Name (e.g., `@Ariel`).
* **Quick Action:** Prominent toggle to switch between **Sub-Region View** and **Super-Region View (Puget Sound)**.

### B. AO Card Component (Workout Sites)

* **Visual Style:** Dark asphalt card with a subtle color indicator strip on the left edge (Blue for Bootcamp, Orange for Ruck, Green for Trail/Specialty).
* **Key Data Points:**
* **AO Name & Site Q:** Bold header (e.g., *The Gasworks*) + Site Q avatar.
* **Time & Days:** Monospace font (e.g., `TUE / THU • 0530`).
* **The Q:** "Q: @Roswell" badge. If open, an active Tac-Orange button: **[ TAKE THE Q ]**.
* **HC Counter:** A quick pill showing total Hard Commits (`14 PAX Committed`).



### C. The "Backblast" & Feed Card

* **Post Cards:** Designed like field reports or tactical logs.
* **Roster Tagging:** Interactive pill badges for every PAX mentioned (e.g., `+12 PAX: @Columbia @Voltaire @KnuckleBuster`).
* **FNG Counter:** Highlighted pill for FNG arrivals.

### D. Leadership & SLT Roster Matrix Widget

* **Grid Layout:** Clean table/card view representing the Regional SLT.
* **Role Badges:** Color-coded status tags:
* **Active Leader** (Confirmed via announcement)
* **Vacant / Recruiting** (High-visibility callout)
* **Emeritus** (Muted dark gray badge)



---

## 6. Micro-Interactions & Tone of Voice

### Haptic Feedback & Audio

* **The "HC" Button Press:** Deep, crisp haptic thud (resembling an iron stamp) when a user hits "Hard Commit" for a workout.
* **Taking the Q:** Satisfying confirmation animation (e.g., shovel flag plant icon locking into place).

### Tone & Microcopy

* **Push Notifications:**
* *Pre-blast:* "Gloom call for tomorrow at Gasworks (0530). Who’s taking the Q?"
* *HC Check-in:* "@Ariel, you HC'd for Gasworks tomorrow. Alarm set?"


* **Empty States:**
* *No Backblast yet:* "No field reports filed for this workout yet. Q, post the backblast!"
* *Unassigned Q:* "No Q assigned. Step up and lead!"



---

## Summary UI Design Token Reference

```json
{
  "color": {
    "background-primary": "#1E2225",
    "surface-card": "#121416",
    "text-primary": "#FFFFFF",
    "text-muted": "#8A949E",
    "brand-accent-orange": "#FF5A36",
    "1st-f-blue": "#3B82F6",
    "2nd-f-amber": "#F59E0B",
    "3rd-f-green": "#10B981"
  },
  "typography": {
    "font-display": "Black Ops One, sans-serif",
    "font-body": "Inter, sans-serif",
    "font-mono": "JetBrains Mono, monospace"
  },
  "border-radius": {
    "card": "8px",
    "button": "4px",
    "badge": "12px"
  }
}

```
