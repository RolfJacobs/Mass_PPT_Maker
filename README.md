# Mass PowerPoint Presentation Generator

Automated tool to create Catholic Mass PowerPoint presentations (`.pptx`) for church projection, using a master template, markdown hymn files, a CSV setup file, and online liturgical information.

---

## Features

- **Dynamic Hymn Population**: Automatically parses hymn markdown files in `assets/`, extracting all verses, author attribution, and the hymn-ending symbol (`_hymn_end.jpg`).
- **Flexible Hymn Counts**: Automatically adjusts slides if fewer hymns are used (e.g., 2 offertory hymns instead of 3, or 1 communion hymn instead of 2).
- **No Consecutive Dividers**: Ensures divider slides (blank slides displaying only the liturgical theme background) never appear consecutively.
- **Online Liturgical Calendar Lookup**: Automatically detects the upcoming Sunday date and retrieves:
  - Liturgical Week Title (e.g., *"27th Sunday in Ordinary Time"*) for **Slide 1**.
  - Responsorial Psalm refrain for the **Response Slide**.
- **Liturgical Theme Divider Background**: Updates all divider slides (and title slide background) with a 4:3 theme image (such as the generated *"Vineyard of the Lord"* background for the 27th Sunday).
- **Preserves Fixed Prayers**: Retains Confessio, Kyrie, Gloria, Credo, Sanctus, Mysterium Fidei, Agnus Dei, and the Prayer to St. Michael from the template.

---

## File Structure

- `generate_mass_ppt.py`: Main Python generation script.
- `mass_template.pptx`: PowerPoint template for the Mass.
- `mass_setup.csv`: Weekly configuration mapping Mass parts to hymn names.
- `ppt_structure.md`: Master structural layout of the Mass presentation.
- `assets/`: Directory containing hymn markdown files (e.g. `the_churchs_one_foundation.md`, `PW/as_the_deer.md`).
- `assets/divider_theme.png`: 4:3 liturgical theme image used as background for divider slides.
- `mass_presentation.pptx`: Generated PowerPoint presentation.

---

## Usage

### Quick Start
To generate the presentation using `mass_setup.csv` and the template:
```bash
python3 generate_mass_ppt.py
```
This produces `mass_presentation.pptx`.

### Command-Line Options

```text
python3 generate_mass_ppt.py [OPTIONS]

Options:
  --csv PATH         Path to CSV setup file (default: mass_setup.csv)
  --template PATH    Path to PPTX template (default: mass_template.pptx)
  --output PATH      Output file path (default: mass_presentation.pptx)
  --date YYYY-MM-DD  Target Sunday date (defaults to upcoming Sunday)
  --title TEXT       Override the title slide text (e.g. "27th Sunday in Ordinary Time")
  --response TEXT    Override the Responsorial Psalm response
  --image PATH       Path to a 4:3 image for divider slides (default: assets/divider_theme.png)
```

### Examples

1. **Generate with custom CSV**:
   ```bash
   python3 generate_mass_ppt.py --csv my_mass_setup.csv --output sunday_mass.pptx
   ```

2. **Generate for a specific date**:
   ```bash
   python3 generate_mass_ppt.py --date 2026-10-11
   ```

3. **Specify custom title, response, and background image**:
   ```bash
   python3 generate_mass_ppt.py \
     --title "28th Sunday in Ordinary Time" \
     --response "I will dwell in the house of the Lord all the days of my life." \
     --image assets/divider_theme.png
   ```

---

## Setup CSV Format

The CSV file (`mass_setup.csv`) maps liturgy sections to hymn file basenames:

```csv
entrance;the_churchs_one_foundation
offertory_1;take_our_bread
offertory_2;to_be_the_body_of_the_lord_in_this_world
offertory_3;christ_be_beside_me
communion_1;as_the_deer
communion_2;all_to_jesus_i_surrender
recessional;go_the_mass_is_ended

mass;mass_of_renew
```

- Sections not present (for example, omitting `offertory_3` or `communion_2`) will be automatically removed from the presentation along with their adjacent divider slides.
- The `mass;mass_of_renew` entry is ignored as Mass prayers are embedded directly in the template.

---

## Web App for Volunteers (WhatsApp & QR Code)

A mobile-friendly web application is available in the [`webapp/`](file:///home/rolf/Documents/Mass_PPT_Maker/webapp/) folder to allow music ministers or volunteers to select weekly hymns on their phone and send the completed setup to you via WhatsApp.

### Features
- **Searchable Autocomplete**: Searches all 280 hymns by title, first line, or filename.
- **Support for New / Uncataloged Hymns**: Type any custom hymn name if it's not yet in the library.
- **Dynamic Offertory & Communion Rows**: Tap `+` or `Remove` to easily choose 1, 2, or 3 hymns.
- **One-Tap WhatsApp Send**: Opens WhatsApp with your phone number and pre-formatted CSV message.
- **Download & Copy**: Also supports directly downloading `mass_setup.csv` or copying to clipboard.
- **Zero Server Overhead**: 100% client-side HTML/JS with embedded catalog; works offline.

### Hosting the Web App (Free on GitHub Pages)
1. Push your repository to GitHub.
2. In the repository settings, go to **Pages**.
3. Under **Build and deployment**, select **Deploy from a branch**, choose `main`, and select `/ (root)` or `/webapp`.
4. Your web app will be live at `https://<your-username>.github.io/<repo>/webapp/`.

### Generating a Printable QR Code
To generate a QR code pointing to your hosted web app:
```bash
python3 generate_qr.py --url "https://<your-username>.github.io/<repo>/webapp/"
```
This generates a high-resolution printable `qr_code.png`.

### Syncing Hymns When You Add New Files
Whenever you add or modify hymn markdown files in `assets/`:
```bash
python3 generate_mass_ppt.py --sync-catalog
```
This re-scans `assets/`, updates `webapp/hymns.json`, and refreshes the embedded catalog in `webapp/index.html`.
