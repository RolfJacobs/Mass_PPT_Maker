---
name: generate-mass-theme
description: >-
  Automatically generates the Catholic liturgical 4:3 theme divider image for the upcoming Sunday Mass and optionally rebuilds the PowerPoint presentation. Trigger when the user mentions "theme image", "make image", "new image", "generate image", "make mass", "update mass", "full mass", or asks for the week's liturgical background picture.
---

# Generate Catholic Mass Theme Image & Presentation

This skill automates creating a Catholic liturgical divider image for the upcoming Sunday Mass and optionally rebuilding `mass_presentation.pptx`.

## Workflow Steps

### Step 1: Determine the Target Sunday & Liturgical Info
Run a quick python check to get the upcoming Sunday date, liturgical title, and Responsorial Psalm:
```bash
python3 -c "
import datetime
from generate_mass_ppt import get_upcoming_sunday, fetch_liturgical_info, get_liturgical_cycle_letter
sunday = get_upcoming_sunday()
title, psalm = fetch_liturgical_info(sunday)
cycle = get_liturgical_cycle_letter(sunday)
print(f'Sunday: {sunday} (Year {cycle})')
print(f'Title: {title}')
print(f'Psalm: {psalm}')
"
```

Look up or query the Gospel and First Reading themes for that Sunday (Year A, B, or C).

### Step 2: Invoke the `image-generator` Subagent
Use the `invoke_subagent` tool with `typeName="image-generator"`, `role="Liturgical Image Designer"`, and a structured prompt following these strict visual guidelines:

1. **Aspect Ratio**: 4:3 (will be used as a divider slide in a 4:3 PowerPoint presentation).
2. **Subject Matter**: A serene, reverent, and artistic biblical depiction capturing the central theme of the Sunday readings (e.g., the Heavenly Banquet, the Good Shepherd, Living Water, etc.).
3. **Lighting & Color**: Soft divine light, cinematic depth, warm golden light and liturgical colors appropriate for the season (Ordinary Time: warm greens and gold; Advent: violet/purple and soft dawn; Lent: violet and solemn desert light; Easter/Christmas: radiant white and gold).
4. **STRICT RULE - NO TEXT**: **NO text, NO words, NO letters, NO typography, NO watermarks** anywhere on the image.
5. **Output Destination**: Save the generated image to `assets/divider_theme.png`. *(Note: If the subagent saves the image in the artifact directory, copy and convert it to `assets/divider_theme.png`)*.

Example subagent invocation prompt:
```text
Generate a 4:3 Catholic liturgical divider slide background image for the [Sunday Title] ([Year Cycle], [Date]).

Theme & Readings:
- Gospel: [Gospel Reference & Theme]
- First Reading: [Reading Reference & Theme]
- Responsorial Psalm: [Psalm Refrain]

Visual Style Guidelines:
- Aspect ratio: 4:3
- A serene, reverent, and artistic biblical depiction of [Theme Details].
- NO text, NO words, NO letters, NO typography on the image.
- High quality, cinematic lighting, warm golden and subtle liturgical tones.

Save the generated image directly to assets/divider_theme.png.
```

### Step 3: Check If Presentation Rebuild Was Requested
If the user requested updating or generating the presentation (e.g. phrases like `"make mass"`, `"update mass"`, `"full mass"`, or if they asked to regenerate):
Run:
```bash
python3 generate_mass_ppt.py
```
*(Note: `generate_mass_ppt.py` automatically updates `webapp/history.json` AND embeds `DEFAULT_MASS_HISTORY` into `webapp/index.html` so offline/local and GitHub Pages deployments both stay up to date without caching issues).*

If hymns or assets were added or changed, also run:
```bash
python3 sync_catalog.py
```

### Step 4: Final Confirmation & Auto-Deploy
1. Verify `assets/divider_theme.png` exists.
2. If `generate_mass_ppt.py` was executed, verify the generated `.pptx` presentation was created, `webapp/history.json` was updated, and `webapp/index.html` contains the embedded entry.
3. Automatically commit and push webapp updates so GitHub Pages deploys seamlessly without manual user intervention:
```bash
git add webapp/ assets/divider_theme.png mass_setup.csv generate_mass_ppt.py sync_catalog.py .agents/
git commit -m "Update mass presentation, history, and theme for [Sunday Title]"
git push origin main
```
4. Present a concise summary to the user:
   - Target Sunday date and liturgical title.
   - Scriptural theme depicted in the image.
   - Clickable link to the updated files: `assets/divider_theme.png` and the generated PowerPoint file (e.g. `[Title] y[Cycle] [Year].pptx`).
   - Confirmation that GitHub Pages has been updated and deployed.
