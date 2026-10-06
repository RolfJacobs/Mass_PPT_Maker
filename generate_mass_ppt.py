#!/usr/bin/env python3
"""
Mass PowerPoint Presentation Generator
Automates generation of Catholic Mass PowerPoint presentations from a template,
a setup CSV, markdown hymn files, and online liturgical information.
"""

import os
import sys
import re
import copy
import argparse
import datetime
import json
import urllib.request
from bs4 import BeautifulSoup
from PIL import Image

from pptx import Presentation
from pptx.parts.presentation import PresentationPart
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Ensure new slide XML parts receive unique partnames even if slides were deleted
PresentationPart._next_slide_partname = property(
    lambda self: self.package.next_partname('/ppt/slides/slide%d.xml')
)


# Mapping of the template slides based on ppt_structure.md
TEMPLATE_SECTIONS = [
    ('title_slide', [0]),
    ('welcome_message', [1]),
    ('divider', [2]),
    ('entrance_hymn', [3, 4, 5]),
    ('divider', [6]),
    ('confessio', [7, 8]),
    ('divider', [9]),
    ('kyrie_eleison', [10]),
    ('divider', [11]),
    ('gloria', [12, 13, 14, 15]),
    ('divider', [16]),
    ('response', [17]),
    ('divider', [18]),
    ('credo', [19, 20, 21, 22, 23, 24, 25]),
    ('divider', [26]),
    ('offertory_1', [27, 28, 29]),
    ('divider', [30]),
    ('offertory_2', [31, 32, 33]),
    ('divider', [34]),
    ('offertory_3', [35, 36, 37]),
    ('divider', [38]),
    ('sanctus', [39, 40]),
    ('divider', [41]),
    ('mysterium_fideii', [42]),
    ('divider', [43]),
    ('agnus_dei', [44, 45]),
    ('divider', [46]),
    ('communion_1', [47, 48, 49]),
    ('divider', [50]),
    ('communion_2', [51, 52, 53]),
    ('divider', [54]),
    ('prayer_to_st_michael', [55, 56]),
    ('divider', [57]),
    ('recessional_hymn', [58, 59, 60]),
    ('divider', [61]),
]

OPTIONAL_HYMN_SECTIONS = {
    'offertory_1': 16, # index in TEMPLATE_SECTIONS
    'offertory_2': 18,
    'offertory_3': 20,
    'communion_1': 28,
    'communion_2': 30,
}

KEY_MAP = {
    'entrance': 'entrance_hymn',
    'entrance_hymn': 'entrance_hymn',
    'offertory_1': 'offertory_1',
    'offertory_2': 'offertory_2',
    'offertory_3': 'offertory_3',
    'communion_1': 'communion_1',
    'communion_2': 'communion_2',
    'recessional': 'recessional_hymn',
    'recessional_hymn': 'recessional_hymn',
}


def get_upcoming_sunday(dt=None):
    """Returns date of upcoming Sunday. If dt is already Sunday, returns dt."""
    if dt is None:
        dt = datetime.date.today()
    days_ahead = (6 - dt.weekday()) % 7
    return dt + datetime.timedelta(days=days_ahead)


def get_liturgical_cycle_letter(target_date):
    """Calculates liturgical cycle (Year A, B, or C) for a given date."""
    year = target_date.year
    # Find 1st Sunday of Advent (Sunday between Nov 27 and Dec 3)
    advent_1 = None
    for d in range(27, 31):
        dt = datetime.date(year, 11, d)
        if dt.weekday() == 6:
            advent_1 = dt
            break
    if not advent_1:
        for d in range(1, 4):
            dt = datetime.date(year, 12, d)
            if dt.weekday() == 6:
                advent_1 = dt
                break
    cycle_year = year + 1 if (advent_1 and target_date >= advent_1) else year
    cycles = {1: 'A', 2: 'B', 0: 'C'}
    return cycles[cycle_year % 3]


def resolve_lectionary_psalm_response(title, target_date=None, event_key=''):
    """
    Looks up the Responsorial Psalm refrain text from webapp/psalms.json.
    Matches the exact dataset and resolution logic used by the webapp.
    """
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'webapp', 'psalms.json')
    if not os.path.exists(json_path):
        json_path = 'webapp/psalms.json'
    if not os.path.exists(json_path):
        return None

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            ps = json.load(f)
    except Exception:
        return None

    y = get_liturgical_cycle_letter(target_date) if target_date else 'A'
    ev_k = (event_key or '').lower()
    nm = (title or '').lower()

    # 1. Ordinary Sunday
    m = re.search(r'ordsunday(\d+)', ev_k) or re.search(r'(\d+)(?:st|nd|rd|th)?\s+sun(?:day)?\s+(?:in|of)\s+ord(?:inary)?\s+time', nm)
    if m:
        w = str(int(m.group(1)))
        return ps.get('ordinary', {}).get(y, {}).get(w)

    # 2. Christ the King (34th Sunday in Ordinary Time)
    if 'christking' in ev_k or 'christ the king' in nm or 'king of the universe' in nm:
        return ps.get('ordinary', {}).get(y, {}).get('34')

    # 3. Trinity
    if 'trinity' in ev_k or 'trinity' in nm:
        return ps.get('solemnities', {}).get('trinity', {}).get(y)

    # 4. Corpus Christi
    if 'corpuschristi' in ev_k or 'corpus christi' in nm or 'body and blood' in nm:
        return ps.get('solemnities', {}).get('corpus_christi', {}).get(y)

    # 5. Holy Family
    if 'holyfamily' in ev_k or 'holy family' in nm:
        return ps.get('solemnities', {}).get('holy_family', {}).get(y)

    # 6. Advent
    m = re.search(r'advent(\d)', ev_k) or re.search(r'(first|second|third|fourth|1st|2nd|3rd|4th|\d+)(?:st|nd|rd|th)?\s+sun(?:day)?\s+of\s+advent', nm)
    if m:
        wMap = {'first': '1', '1st': '1', 'second': '2', '2nd': '2', 'third': '3', '3rd': '3', 'fourth': '4', '4th': '4'}
        w = wMap.get(m.group(1).lower(), m.group(1))
        return ps.get('advent', {}).get(y, {}).get(w)

    # 7. Lent
    if 'palmsun' in ev_k or 'palm sunday' in nm or 'palm sun' in nm:
        return ps.get('lent', {}).get(y, {}).get('palm')
    m = re.search(r'lent(\d)', ev_k) or re.search(r'(first|second|third|fourth|fifth|1st|2nd|3rd|4th|5th|\d+)(?:st|nd|rd|th)?\s+sun(?:day)?\s+of\s+lent', nm)
    if m:
        wMap = {'first': '1', '1st': '1', 'second': '2', '2nd': '2', 'third': '3', '3rd': '3', 'fourth': '4', '4th': '4', 'fifth': '5', '5th': '5'}
        w = wMap.get(m.group(1).lower(), m.group(1))
        return ps.get('lent', {}).get(y, {}).get(w)

    # 8. Easter
    if ev_k == 'easter' or 'easter sunday' in nm or 'easter sun' in nm or nm == 'easter':
        return ps.get('easter', {}).get(y, {}).get('1')
    if 'ascension' in ev_k or 'ascension' in nm:
        return ps.get('easter', {}).get(y, {}).get('ascension')
    if 'pentecost' in ev_k or 'pentecost' in nm:
        return ps.get('easter', {}).get(y, {}).get('pentecost')
    m = re.search(r'easter(\d)', ev_k) or re.search(r'(second|third|fourth|fifth|sixth|seventh|2nd|3rd|4th|5th|6th|7th|\d+)(?:st|nd|rd|th)?\s+sun(?:day)?\s+of\s+easter', nm)
    if m:
        wMap = {'second': '2', '2nd': '2', 'third': '3', '3rd': '3', 'fourth': '4', '4th': '4', 'fifth': '5', '5th': '5', 'sixth': '6', '6th': '6', 'seventh': '7', '7th': '7'}
        w = wMap.get(m.group(1).lower(), m.group(1))
        return ps.get('easter', {}).get(y, {}).get(w)

    # Solemnities and feasts
    for s_key in ['christmas2', 'epiphany', 'baptism', 'all_saints']:
        if s_key in ev_k or s_key.replace('_', ' ') in nm or s_key in nm:
            return ps.get('solemnities', {}).get(s_key)

    return None


def fetch_liturgical_info(target_date):
    """
    Fetches the Catholic liturgical title and Responsorial Psalm response
    for target_date (Roman Rite).
    Uses the shared lectionary dataset first, and falls back to CatholicReadings.org.
    """
    year, month, day = target_date.year, target_date.month, target_date.day
    title = None
    event_key = ''
    response_text = None

    # 1. Fetch celebration title (try LitCal first to match webapp, then Calapi)
    litcal_url = f"https://litcal.johnromanodorazio.com/api/v5/calendar?year={year}&locale=en"
    try:
        req = urllib.request.Request(litcal_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=4) as resp:
            data = json.loads(resp.read().decode())
            events = data.get('litcal', [])
            date_prefix = f"{year}-{month:02d}-{day:02d}"
            ev = next((e for e in events if e.get('date', '').startswith(date_prefix)), None)
            if ev:
                title = ev.get('name')
                event_key = ev.get('event_key', '')
    except Exception:
        pass

    if not title:
        cal_url = f"http://calapi.inadiutorium.cz/api/v0/en/calendars/default/{year}/{month:02d}/{day:02d}"
        try:
            req = urllib.request.Request(cal_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=4) as resp:
                data = json.loads(resp.read().decode())
                if data.get('celebrations'):
                    title = data['celebrations'][0].get('title')
        except Exception as e:
            print(f"[Warning] Failed to fetch calendar from calapi: {e}")

    # Fallback title if None
    if not title:
        title = "27th Sunday in Ordinary Time"

    # 2. Responsorial Psalm Refrain from shared lectionary database (webapp/psalms.json)
    response_text = resolve_lectionary_psalm_response(title, target_date, event_key)

    # 3. If not in local dataset, fall back to online CatholicReadings.org scrape
    if not response_text:
        ord_map = [
            (r'\b33rd\b', 'thirty-third'), (r'\b32nd\b', 'thirty-second'), (r'\b31st\b', 'thirty-first'),
            (r'\b30th\b', 'thirtieth'), (r'\b29th\b', 'twenty-ninth'), (r'\b28th\b', 'twenty-eighth'),
            (r'\b27th\b', 'twenty-seventh'), (r'\b26th\b', 'twenty-sixth'), (r'\b25th\b', 'twenty-fifth'),
            (r'\b24th\b', 'twenty-fourth'), (r'\b23rd\b', 'twenty-third'), (r'\b22nd\b', 'twenty-second'),
            (r'\b21st\b', 'twenty-first'), (r'\b20th\b', 'twentieth'), (r'\b19th\b', 'nineteenth'),
            (r'\b18th\b', 'eighteenth'), (r'\b17th\b', 'seventeenth'), (r'\b16th\b', 'sixteenth'),
            (r'\b15th\b', 'fifteenth'), (r'\b14th\b', 'fourteenth'), (r'\b13th\b', 'thirteenth'),
            (r'\b12th\b', 'twelfth'), (r'\b11th\b', 'eleventh'), (r'\b10th\b', 'tenth'),
            (r'\b9th\b', 'ninth'), (r'\b8th\b', 'eighth'), (r'\b7th\b', 'seventh'),
            (r'\b6th\b', 'sixth'), (r'\b5th\b', 'fifth'), (r'\b4th\b', 'fourth'),
            (r'\b3rd\b', 'third'), (r'\b2nd\b', 'second'), (r'\b1st\b', 'first')
        ]
        clean_title = title.lower()
        for pat, word in ord_map:
            clean_title = re.sub(pat, word, clean_title)
        slug = re.sub(r'[^a-z0-9]+', '-', clean_title).strip('-')

        slug_candidates = [
            f"{slug}-year-a",
            f"{slug}-year-b",
            f"{slug}-year-c",
            slug
        ]

        import gzip
        for sc in slug_candidates:
            url = f"https://catholicreadings.org/{sc}/"
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
                with urllib.request.urlopen(req, timeout=6) as resp:
                    raw_bytes = resp.read()
                    if raw_bytes[:2] == b'\x1f\x8b':
                        raw_bytes = gzip.decompress(raw_bytes)
                    soup = BeautifulSoup(raw_bytes.decode('utf-8', errors='ignore'), 'html.parser')
                    for tag in soup.find_all(['strong', 'p']):
                        t = tag.get_text().strip()
                        if t.startswith('R.') or t.startswith('R '):
                            cleaned = re.sub(r'^R\.?\s*(?:\([^)]*\))?\s*', '', t).strip()
                            first_line = cleaned.splitlines()[0].strip()
                            if len(first_line) > 5:
                                response_text = first_line
                                break
                    if response_text:
                        break
            except Exception:
                pass

    return title, response_text


def format_psalm_response(text):
    """Formats psalm response text into balanced uppercase lines for slide display."""
    text = text.strip().rstrip('.').upper()
    words = text.split()
    if len(text) <= 28 or len(words) <= 3:
        return [text]

    # Split into balanced lines
    mid = len(words) // 2
    # Adjust split point to avoid orphan prepositions if possible
    if mid > 1 and words[mid] in ['IN', 'OF', 'TO', 'AND', 'FOR', 'ON', 'WITH']:
        pass
    line1 = " ".join(words[:mid])
    line2 = " ".join(words[mid:])
    return [line1, line2]


def parse_setup_csv(csv_path_or_text):
    """
    Parses mass_setup.csv (or pasted WhatsApp text message) into a dictionary of {section_name: hymn_name}.
    Ignores mass;mass_of_renew or empty entries.
    Supports arbitrary numbers of offertory and communion hymns.
    """
    hymns = {}
    if os.path.isfile(csv_path_or_text):
        with open(csv_path_or_text, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    else:
        lines = csv_path_or_text.splitlines()

    for line in lines:
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        parts = [p.strip() for p in re.split(r'[;,]', line) if p.strip()]
        if len(parts) >= 2:
            key, val = parts[0].lower(), parts[1]
            if key == 'mass':
                hymns['mass_setting'] = val
                continue
            if key in ('response', 'psalm', 'psalm_response'):
                hymns['response'] = val
                continue
            if key in ('title', 'celebration', 'celebration_title'):
                hymns['celebration_title'] = val
                continue
            if key in ('entrance', 'entrance_hymn'):
                hymns['entrance_hymn'] = val
            elif key in ('recessional', 'recessional_hymn'):
                hymns['recessional_hymn'] = val
            elif key == 'offertory':
                if len(parts) > 2:
                    for i, p in enumerate(parts[1:], start=1):
                        if p:
                            hymns[f'offertory_{i}'] = p
                else:
                    hymns['offertory_1'] = val
            elif key == 'communion':
                if len(parts) > 2:
                    for i, p in enumerate(parts[1:], start=1):
                        if p:
                            hymns[f'communion_{i}'] = p
                else:
                    hymns['communion_1'] = val
            elif re.match(r'^offertory_\d+$', key):
                hymns[key] = val
            elif re.match(r'^communion_\d+$', key):
                hymns[key] = val
            elif key in KEY_MAP:
                hymns[KEY_MAP[key]] = val
    return hymns


def resolve_hymn_file(hymn_name, assets_dir='assets'):
    """Finds the markdown file for a given hymn name in assets/ or subfolders."""
    name = hymn_name.strip()
    if not name.endswith('.md'):
        name += '.md'

    direct = os.path.join(assets_dir, name)
    if os.path.exists(direct):
        return direct

    for root, dirs, files in os.walk(assets_dir):
        for f in files:
            if f.lower() == name.lower():
                return os.path.join(root, f)

    return None


def parse_hymn_markdown(filepath):
    """
    Parses a hymn markdown file into sections (verses) and author.
    Sections are separated by '---'.
    """
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    raw_sections = re.split(r'\n\s*---\s*\n', content)
    sections = []
    author = None

    for s in raw_sections:
        s = s.strip()
        if not s:
            continue

        # Extract author from <span ...>-Author</span>
        span_match = re.search(r'<span[^>]*>(.*?)</span>', s, re.IGNORECASE | re.DOTALL)
        if span_match:
            auth_raw = span_match.group(1).strip()
            if auth_raw.startswith('-'):
                auth_raw = auth_raw[1:].strip()
            author = auth_raw
            s = re.sub(r'<span[^>]*>.*?</span>', '', s, flags=re.IGNORECASE | re.DOTALL)

        # Remove markdown image links
        s = re.sub(r'!\[.*?\]\(.*?\)', '', s)

        # Split into non-empty lines, ignoring markdown titles/headings
        lines = [line.strip() for line in s.splitlines() if line.strip() and not line.strip().startswith('#')]
        if lines:
            sections.append(lines)

    return sections, author


def get_hymn_shapes(slide):
    """Identifies header_shape, lyrics_shape, pic_shape on a hymn slide."""
    header_shape = None
    lyrics_shape = None
    pic_shape = None

    for s in slide.shapes:
        if s.shape_type == 13: # Picture
            pic_shape = s
        elif s.has_text_frame:
            t = s.text_frame.text.strip().lower()
            if t in ['entrance', 'offertory', 'communion', 'recessional'] or (s.top < 600000 and s.height < 1000000 and ('title' in s.name.lower() or 'text box' in s.name.lower())):
                header_shape = s
            else:
                lyrics_shape = s

    # Fallback if both text frames were assigned
    if not lyrics_shape and header_shape:
        lyrics_shape = header_shape
        header_shape = None

    return header_shape, lyrics_shape, pic_shape


def populate_lyrics(lyrics_shape, lines, author=None):
    """Populates lines of lyrics and optional author into a text frame with clean styling."""
    tf = lyrics_shape.text_frame
    tf.clear()

    # Determine optimal font size based on text density
    num_lines = len(lines)
    max_len = max(len(l) for l in lines) if lines else 0

    if num_lines >= 8 or max_len > 38:
        font_sz = Pt(22)
    elif num_lines >= 6 or max_len > 32:
        font_sz = Pt(25)
    else:
        font_sz = Pt(28)

    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = PP_ALIGN.CENTER
        p.font.name = 'Arial'
        p.font.size = font_sz
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

    if author:
        # Blank line
        p_blank = tf.add_paragraph()
        p_blank.text = ""
        p_blank.font.size = Pt(10)

        p_auth = tf.add_paragraph()
        p_auth.text = author if author.startswith('-') else f"-{author}"
        p_auth.alignment = PP_ALIGN.CENTER
        p_auth.font.name = 'Arial'
        p_auth.font.size = Pt(18)
        p_auth.font.bold = False
        p_auth.font.color.rgb = RGBColor(255, 242, 204) # Gold/cream


def clone_slide_to_position(prs, source_slide, insert_before_idx):
    """Clones source_slide and inserts it at insert_before_idx in presentation."""
    dest_slide = prs.slides.add_slide(source_slide.slide_layout)

    # Clear generated placeholders
    spTree = dest_slide.shapes._spTree
    for sp in list(spTree)[2:]:
        spTree.remove(sp)

    # Copy relationships
    rId_map = {}
    for rel_id, rel in source_slide.part.rels.items():
        if rel.reltype.endswith('image'):
            new_rid = dest_slide.part.relate_to(rel.target_part, rel.reltype)
            rId_map[rel_id] = new_rid

    # Deep copy shapes
    src_spTree = source_slide.shapes._spTree
    for sp in list(src_spTree)[2:]:
        new_sp = copy.deepcopy(sp)
        for old_rid, new_rid in rId_map.items():
            for elem in new_sp.xpath(f'.//*[@*="{old_rid}"]'):
                for attr in elem.attrib:
                    if elem.attrib[attr] == old_rid:
                        elem.attrib[attr] = new_rid
        spTree.append(new_sp)

    # Move slide ID in presentation to insert_before_idx
    new_sldId = prs.slides._sldIdLst[-1]
    prs.slides._sldIdLst.insert(insert_before_idx, new_sldId)

    return prs.slides[insert_before_idx]


def delete_slides(prs, indices):
    """Deletes slides by presentation index in descending order."""
    for idx in sorted(indices, reverse=True):
        r_id = prs.slides._sldIdLst[idx].rId
        prs.part.drop_rel(r_id)
        del prs.slides._sldIdLst[idx]


def is_divider_slide(slide, slide_index):
    """Determines whether a slide is a divider slide (only background, no text)."""
    if slide_index == 0: # Slide 1 is title slide
        return False
    for shape in slide.shapes:
        if shape.shape_type == 6: # Group
            for sub in shape.shapes:
                if sub.has_text_frame and sub.text_frame.text.strip():
                    return False
        elif shape.has_text_frame and shape.text_frame.text.strip():
            return False
    return True


def remove_adjacent_dividers(prs):
    """Ensures there are NEVER two divider slides next to each other."""
    removed = 0
    i = 0
    while i < len(prs.slides) - 1:
        if is_divider_slide(prs.slides[i], i) and is_divider_slide(prs.slides[i+1], i+1):
            r_id = prs.slides._sldIdLst[i+1].rId
            prs.part.drop_rel(r_id)
            del prs.slides._sldIdLst[i+1]
            removed += 1
        else:
            i += 1
    return removed


def update_divider_background(prs, image_path):
    """
    Replaces the shared divider slide background image (image3.png) with image_path.
    This automatically updates all divider slides and Slide 1 background.
    """
    if not os.path.exists(image_path):
        print(f"[Warning] Background image {image_path} not found. Skipping background update.")
        return

    # Convert/resize image to PNG with 4:3 ratio if needed
    im = Image.open(image_path)
    w, h = im.size
    target_ratio = 4.0 / 3.0
    current_ratio = w / h
    if abs(current_ratio - target_ratio) > 0.02:
        if current_ratio > target_ratio:
            new_w = int(h * target_ratio)
            left = (w - new_w) // 2
            im = im.crop((left, 0, left + new_w, h))
        else:
            new_h = int(w / target_ratio)
            top = (h - new_h) // 2
            im = im.crop((0, top, w, top + new_h))

    # Resize to 1024x768 (standard 4:3 projection)
    im = im.resize((1024, 768), Image.Resampling.LANCZOS)
    import io
    buf = io.BytesIO()
    im.save(buf, format='PNG')
    png_bytes = buf.getvalue()

    # Find image3.png part from slide 3 (divider)
    slide3 = prs.slides[2]
    blip = slide3.background._element.xpath('.//a:blip/@r:embed')[0]
    img_part = slide3.part.rels[blip].target_part
    img_part._blob = png_bytes
    print(f"[Success] Updated divider background using {image_path} (1024x768 4:3)")


def shorten_celebration_title(title):
    """Shortens 'Sunday' to 'Sun' and 'Ordinary' to 'Ord' for neat layout wrapping."""
    if not title:
        return title
    title = re.sub(r'\bSunday\b', 'Sun', title, flags=re.IGNORECASE)
    title = re.sub(r'\bOrdinary\b', 'Ord', title, flags=re.IGNORECASE)
    return title


def update_title_slide(prs, title_text):
    """Updates Slide 1 with Catholic Church week title (e.g. '27th Sun in Ord Time')."""
    slide1 = prs.slides[0]
    # Shape 0 is GROUP containing banner picture and text box
    grp = slide1.shapes[0]
    txt_box = grp.shapes[1]
    tf = txt_box.text_frame
    shortened_title = shorten_celebration_title(title_text)
    tf.text = shortened_title
    for p in tf.paragraphs:
        p.font.name = 'Calibri'
        p.font.size = Pt(32)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
    print(f"[Success] Title slide updated: '{shortened_title}'")


def enforce_center_alignment(prs):
    """Rule to ensure all text across all shapes, groups, and slides is centre-aligned."""
    def align_shape(shape):
        if shape.shape_type == 6:  # GROUP shape
            for subshape in shape.shapes:
                align_shape(subshape)
        elif shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER

    for slide in prs.slides:
        for shape in slide.shapes:
            align_shape(shape)


def update_response_slide(prs, response_lines):
    """Updates Slide 18 with Responsorial Psalm refrain."""
    # Find response slide (Slide 18 in template)
    slide18 = prs.slides[17]
    sub = slide18.shapes[0]
    tf = sub.text_frame
    tf.clear()

    for i, line in enumerate(response_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = PP_ALIGN.CENTER
        p.font.name = 'Arial'
        p.font.size = Pt(38) if len(line) <= 24 else Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

    print(f"[Success] Response slide updated: {' / '.join(response_lines)}")


def expand_template_sections(prs, hymn_setup):
    """
    Dynamically expands template slides and section mappings if hymn_setup contains
    more than 3 offertory hymns or more than 2 communion hymns.
    Clones slides from existing prototype sections and inserts them with divider slides.
    Returns (expanded_template_sections, optional_hymn_sections).
    """
    sections = [list(item) for item in copy.deepcopy(TEMPLATE_SECTIONS)]

    # Prototypes from initial template:
    # offertory_3: slides 35, 36, 37; divider: slide 38
    # communion_2: slides 51, 52, 53; divider: slide 54
    offertory_proto = [prs.slides[35], prs.slides[36], prs.slides[37]]
    off_divider_proto = prs.slides[38]
    communion_proto = [prs.slides[51], prs.slides[52], prs.slides[53]]
    comm_divider_proto = prs.slides[54]

    # Check max offertory index in setup
    off_nums = [int(m.group(1)) for k in hymn_setup for m in [re.match(r'^offertory_(\d+)$', k)] if m]
    max_off = max(off_nums) if off_nums else 3

    if max_off > 3:
        for off_num in range(4, max_off + 1):
            sanctus_pos = [s[0] for s in sections].index('sanctus')
            insert_pos = sections[sanctus_pos][1][0]
            clone_slide_to_position(prs, offertory_proto[0], insert_pos)
            clone_slide_to_position(prs, offertory_proto[1], insert_pos + 1)
            clone_slide_to_position(prs, offertory_proto[2], insert_pos + 2)
            clone_slide_to_position(prs, off_divider_proto, insert_pos + 3)
            sections.insert(sanctus_pos, [f'offertory_{off_num}', [insert_pos, insert_pos + 1, insert_pos + 2]])
            sections.insert(sanctus_pos + 1, ['divider', [insert_pos + 3]])
            for s in sections[sanctus_pos + 2:]:
                s[1] = [idx + 4 for idx in s[1]]
            print(f"[Expand] Added template slides for offertory_{off_num} (slides {insert_pos}..{insert_pos+3})")

    # Check max communion index in setup
    comm_nums = [int(m.group(1)) for k in hymn_setup for m in [re.match(r'^communion_(\d+)$', k)] if m]
    max_comm = max(comm_nums) if comm_nums else 2

    if max_comm > 2:
        for comm_num in range(3, max_comm + 1):
            michael_pos = [s[0] for s in sections].index('prayer_to_st_michael')
            insert_pos = sections[michael_pos][1][0]
            clone_slide_to_position(prs, communion_proto[0], insert_pos)
            clone_slide_to_position(prs, communion_proto[1], insert_pos + 1)
            clone_slide_to_position(prs, communion_proto[2], insert_pos + 2)
            clone_slide_to_position(prs, comm_divider_proto, insert_pos + 3)
            sections.insert(michael_pos, [f'communion_{comm_num}', [insert_pos, insert_pos + 1, insert_pos + 2]])
            sections.insert(michael_pos + 1, ['divider', [insert_pos + 3]])
            for s in sections[michael_pos + 2:]:
                s[1] = [idx + 4 for idx in s[1]]
            print(f"[Expand] Added template slides for communion_{comm_num} (slides {insert_pos}..{insert_pos+3})")

    optional_hymn_sections = {
        s[0] for s in sections
        if s[0].startswith('offertory_') or s[0].startswith('communion_')
    }

    return [(item[0], item[1]) for item in sections], optional_hymn_sections


MASS_PARTS = {
    'kyrie_eleison': '1_kyrie.md',
    'gloria': '2_gloria.md',
    'sanctus': '3_sanctus.md',
    'mysterium_fideii': '4_mysterium_fideii.md',
    'agnus_dei': '5_agnus_dei.md',
}

MASS_FOLDER_MAP = {
    'mass_of_renew': 'Mass_of_Renew',
    'mass_of_renewal': 'Mass_of_Renew',
    'renew': 'Mass_of_Renew',
    'christ_the_saviour': 'Christ_the_Saviour',
    'mass_of_christ_the_saviour': 'Christ_the_Saviour',
    'st_john': 'St_John',
    'mass_of_st_john': 'St_John',
    'st_thomas_moore': 'St_Thomas_Moore',
    'mass_of_st_thomas_moore': 'St_Thomas_Moore',
    'st_thomas_more': 'St_Thomas_Moore',
    'mass_of_st_thomas_more': 'St_Thomas_Moore',
}


def resolve_mass_setting(setting_name):
    """Resolves mass setting name from CSV to folder name in assets/Masses/."""
    if not setting_name:
        return 'Mass_of_Renew'
    clean = re.sub(r'[^a-z0-9]+', '_', setting_name.strip().lower()).strip('_')
    if clean in MASS_FOLDER_MAP:
        return MASS_FOLDER_MAP[clean]
    for k, v in MASS_FOLDER_MAP.items():
        if k in clean or clean in k:
            return v
    return 'Mass_of_Renew'


def populate_mass_part(prs, sec_name, slide_indices, md_file_path):
    """Populates mass part slides from markdown, cloning or deleting slides as needed."""
    if not os.path.exists(md_file_path):
        print(f"[Warning] Mass part file '{md_file_path}' not found.")
        return

    verses, author = parse_hymn_markdown(md_file_path)
    if not verses:
        print(f"[Warning] No verses parsed from '{md_file_path}'.")
        return

    N = len(verses)
    M = len(slide_indices)
    print(f"[Populate Mass] Section '{sec_name}' from '{md_file_path}' ({N} verses, author: {repr(author)})")

    if M == 1 and N > 1:
        _, l0, _ = get_hymn_shapes(prs.slides[slide_indices[0]])
        populate_lyrics(l0, verses[0])
        insert_pos = slide_indices[0] + 1
        template_to_clone = prs.slides[slide_indices[0]]
        for v_i in range(1, N):
            new_slide = clone_slide_to_position(prs, template_to_clone, insert_pos)
            h_new, l_new, _ = get_hymn_shapes(new_slide)
            if h_new and h_new != l_new:
                h_new.text_frame.clear()
            is_last = (v_i == N - 1)
            populate_lyrics(l_new, verses[v_i], author=author if is_last else None)
            insert_pos += 1
        return

    if N <= M:
        for i in range(N - 1):
            _, l, _ = get_hymn_shapes(prs.slides[slide_indices[i]])
            populate_lyrics(l, verses[i])
        _, l_last, _ = get_hymn_shapes(prs.slides[slide_indices[N - 1]])
        populate_lyrics(l_last, verses[N - 1], author=author)
        if M > N:
            delete_slides(prs, slide_indices[N:])
    else:
        for i in range(M - 1):
            _, l, _ = get_hymn_shapes(prs.slides[slide_indices[i]])
            populate_lyrics(l, verses[i])
        last_slide_idx = slide_indices[-1]
        template_to_clone = prs.slides[slide_indices[-1]]
        for v_i in range(M - 1, N - 1):
            new_slide = clone_slide_to_position(prs, template_to_clone, last_slide_idx)
            h_new, l_new, _ = get_hymn_shapes(new_slide)
            if h_new and h_new != l_new:
                h_new.text_frame.clear()
            populate_lyrics(l_new, verses[v_i])
            last_slide_idx += 1
        _, l_last, _ = get_hymn_shapes(prs.slides[last_slide_idx])
        populate_lyrics(l_last, verses[-1], author=author)


def generate_presentation(csv_path, template_path, output_path, title_text=None, response_text=None, image_path=None, target_date=None):
    """Main generation workflow."""
    print("=" * 60)
    print("Mass PowerPoint Presentation Generator")
    print("=" * 60)

    # 1. Liturgical Date & Online Lookup
    if target_date is None:
        target_date = get_upcoming_sunday()
    print(f"Target Liturgical Date: {target_date} (Sunday)")

    if not title_text or not response_text:
        fetched_title, fetched_resp = fetch_liturgical_info(target_date)
        if not title_text and fetched_title:
            title_text = fetched_title
        if not response_text and fetched_resp:
            response_text = fetched_resp

    if not title_text:
        title_text = "27th Sunday in Ordinary Time"
    if not response_text:
        response_text = "The vineyard of the Lord is the house of Israel."

    print(f"Liturgical Title: {title_text}")
    print(f"Psalm Response Refrain: {response_text}")

    # 2. Parse CSV
    print(f"Loading hymn setup from: {csv_path}")
    hymn_setup = parse_setup_csv(csv_path)
    for k, v in hymn_setup.items():
        print(f"  {k}: {v}")

    # Allow CSV / setup message to supply response refrain or celebration title
    if hymn_setup.get('response'):
        response_text = hymn_setup['response']
        print(f"Overriding Psalm Response Refrain from CSV: {response_text}")
    if hymn_setup.get('celebration_title'):
        title_text = hymn_setup['celebration_title']
        print(f"Overriding Liturgical Title from CSV: {title_text}")

    raw_mass = hymn_setup.get('mass_setting', 'Mass of Renew')
    mass_folder = resolve_mass_setting(raw_mass)
    print(f"Selected Mass Setting: {mass_folder} (input: '{raw_mass}')")

    # 3. Load Presentation Template
    print(f"Loading template: {template_path}")
    prs = Presentation(template_path)

    # 4. Update Divider Background Image
    if image_path and os.path.exists(image_path):
        update_divider_background(prs, image_path)
    elif os.path.exists("assets/divider_theme.png"):
        update_divider_background(prs, "assets/divider_theme.png")
    else:
        print("[Info] No custom divider image found; keeping template background.")

    # 5. Update Title Slide & Response Slide
    update_title_slide(prs, title_text)
    update_response_slide(prs, format_psalm_response(response_text))

    # 6. Process Sections (Backwards to preserve slide index stability)
    template_sections, optional_hymn_sections = expand_template_sections(prs, hymn_setup)
    sections_reversed = list(reversed(template_sections))

    for sec_name, slide_indices in sections_reversed:
        # Check if this is one of the 5 mass parts
        if sec_name in MASS_PARTS:
            md_path = os.path.join('assets', 'Masses', mass_folder, MASS_PARTS[sec_name])
            populate_mass_part(prs, sec_name, slide_indices, md_path)
            continue

        if 'hymn' not in sec_name and 'offertory' not in sec_name and 'communion' not in sec_name:
            continue

        is_optional = sec_name in optional_hymn_sections
        hymn_key = hymn_setup.get(sec_name)

        if not hymn_key:
            if is_optional:
                # Remove this hymn section AND the divider slide following it
                # Find divider slide after this section in template_sections
                sec_pos = [s[0] for s in template_sections].index(sec_name)
                next_sec_name, next_sec_indices = template_sections[sec_pos + 1]
                to_delete = list(slide_indices)
                if next_sec_name == 'divider':
                    to_delete.extend(next_sec_indices)
                print(f"[Remove] Section '{sec_name}' omitted from CSV. Deleting slides {to_delete}...")
                delete_slides(prs, to_delete)
            continue

        # Hymn is present: resolve markdown file
        hymn_path = resolve_hymn_file(hymn_key)
        if not hymn_path:
            print(f"[Warning] Hymn file for '{hymn_key}' not found in assets. Leaving placeholder.")
            continue

        verses, author = parse_hymn_markdown(hymn_path)
        if not verses:
            print(f"[Warning] No verses parsed from '{hymn_path}'.")
            continue

        print(f"[Populate] Section '{sec_name}' with '{hymn_key}' ({len(verses)} verses, author: {repr(author)})")

        N = len(verses)
        s0_idx, s1_idx, s2_idx = slide_indices[0], slide_indices[1], slide_indices[2]
        slide0 = prs.slides[s0_idx]
        slide1 = prs.slides[s1_idx]
        slide2 = prs.slides[s2_idx]

        h0, l0, p0 = get_hymn_shapes(slide0)
        h1, l1, p1 = get_hymn_shapes(slide1)
        h2, l2, p2 = get_hymn_shapes(slide2)

        if N >= 3:
            # Verse 1 -> slide0
            populate_lyrics(l0, verses[0])
            # Verse 2 -> slide1
            populate_lyrics(l1, verses[1])

            # If N > 3, clone slide1 for verses 2 through N-2
            if N > 3:
                # We need to insert before slide2 (current s2_idx)
                # Each insertion shifts s2_idx by 1, so inserting at s2_idx repeatedly places them in order!
                for v_i in range(2, N - 1):
                    new_slide = clone_slide_to_position(prs, slide1, s2_idx)
                    _, new_l, _ = get_hymn_shapes(new_slide)
                    populate_lyrics(new_l, verses[v_i])
                    s2_idx += 1 # update slide2 index

            # Last verse -> slide2 with author
            populate_lyrics(l2, verses[-1], author=author)

        elif N == 2:
            populate_lyrics(l0, verses[0])
            populate_lyrics(l2, verses[1], author=author)
            # Delete middle slide
            delete_slides(prs, [s1_idx])

        elif N == 1:
            populate_lyrics(l0, verses[0], author=author)
            # Add _hymn_end.jpg to slide0
            if os.path.exists("assets/_assets/_hymn_end.jpg"):
                slide0.shapes.add_picture("assets/_assets/_hymn_end.jpg", 3962400, 5029200, 1219200, 1544637)
            # Delete middle and last slide
            delete_slides(prs, [s1_idx, s2_idx])

    # 7. Remove adjacent dividers (guarantee rule)
    removed_adj = remove_adjacent_dividers(prs)
    if removed_adj > 0:
        print(f"[Sanity Check] Removed {removed_adj} adjacent divider slide(s).")

    # 8. Enforce rule: all text across all slides is centre aligned
    enforce_center_alignment(prs)

    # 9. Save Final Presentation
    prs.save(output_path)
    print(f"\n[Done] Successfully generated presentation: {output_path} (Total slides: {len(prs.slides)})")
    print("=" * 60)


def main():
    parser = argparse.ArgumentParser(description="Automate Catholic Mass PowerPoint generation.")
    parser.add_argument("--csv", default="mass_setup.csv", help="Path to input CSV setup file.")
    parser.add_argument("--template", default="mass_template.pptx", help="Path to PowerPoint template.")
    parser.add_argument("--output", default="mass_presentation.pptx", help="Path for generated output presentation.")
    parser.add_argument("--title", default=None, help="Override title slide text (e.g. '27th Sunday in Ordinary Time').")
    parser.add_argument("--response", default=None, help="Override responsorial psalm response text.")
    parser.add_argument("--image", default=None, help="Path to 4:3 background image for divider slides.")
    parser.add_argument("--date", default=None, help="Target date in YYYY-MM-DD format (defaults to upcoming Sunday).")
    parser.add_argument("--sync-catalog", action="store_true", help="Sync assets catalog to webapp/hymns.json.")

    args = parser.parse_args()

    if args.sync_catalog:
        from sync_catalog import sync_catalog
        sync_catalog()
        return

    target_date = None
    if args.date:
        try:
            target_date = datetime.date.fromisoformat(args.date)
        except ValueError:
            print(f"[Error] Invalid date format '{args.date}'. Expected YYYY-MM-DD.")
            sys.exit(1)

    generate_presentation(
        csv_path=args.csv,
        template_path=args.template,
        output_path=args.output,
        title_text=args.title,
        response_text=args.response,
        image_path=args.image,
        target_date=target_date
    )


if __name__ == "__main__":
    main()
