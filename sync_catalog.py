#!/usr/bin/env python3
"""
Hymn Catalog Synchronizer
Scans the assets/ directory for hymn markdown files and generates webapp/hymns.json
and embeds the catalog into webapp/index.html so it works standalone/offline.
"""

import os
import re
import json

def clean_title_from_filename(fname):
    base = os.path.splitext(os.path.basename(fname))[0]
    words = base.split('_')
    title = ' '.join(w.capitalize() for w in words)
    title = title.replace('Pw', '(P&W)').replace('The Churchs', "The Church's").replace('Lords', "Lord's")
    return title

def sync_catalog(assets_dir='assets', webapp_dir='webapp'):
    os.makedirs(webapp_dir, exist_ok=True)
    hymns = []

    for root, dirs, files in os.walk(assets_dir):
        if 'Masses' in root or 'Prayers' in root or 'testing' in root:
            continue
        for f in sorted(files):
            if not f.endswith('.md'):
                continue
            filepath = os.path.join(root, f)
            base_id = os.path.splitext(f)[0]
            category = 'Praise & Worship' if 'PW' in root else 'Traditional'

            author = None
            first_line = ''
            try:
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as fp:
                    content = fp.read()
                    m = re.search(r'<span[^>]*>(.*?)</span>', content, re.IGNORECASE)
                    if m:
                        auth = m.group(1).strip()
                        if auth.startswith('-'): auth = auth[1:].strip()
                        author = auth
                    for line in content.splitlines():
                        l = line.strip()
                        if l and not l.startswith('---') and not l.startswith('<') and not l.startswith('!['):
                            first_line = l
                            break
            except Exception:
                pass

            title = clean_title_from_filename(f)
            hymns.append({
                'id': base_id,
                'title': title,
                'first_line': first_line,
                'author': author or '',
                'category': category
            })

    json_path = os.path.join(webapp_dir, 'hymns.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(hymns, f, indent=2)

    # Also embed directly in webapp/index.html as fallback for zero-network / local file opening
    html_path = os.path.join(webapp_dir, 'index.html')
    if os.path.exists(html_path):
        with open(html_path, 'r', encoding='utf-8') as f:
            html = f.read()
        embedded_json = json.dumps(hymns)
        updated_html = re.sub(
            r'let HYMN_CATALOG = \[.*?\];',
            lambda m: f'let HYMN_CATALOG = {embedded_json};',
            html,
            flags=re.DOTALL
        )
        # Also sync psalms.json if available
        psalms_path = os.path.join(webapp_dir, 'psalms.json')
        if os.path.exists(psalms_path):
            with open(psalms_path, 'r', encoding='utf-8') as pf:
                psalms_data = json.load(pf)
            psalms_json = json.dumps(psalms_data, separators=(',', ':'))
            updated_html = re.sub(
                r'let LECTIONARY_PSALM_RESPONSES = \{.*?\};',
                lambda m: f'let LECTIONARY_PSALM_RESPONSES = {psalms_json};',
                updated_html,
                flags=re.DOTALL
            )
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(updated_html)
        print(f"[Success] Catalog & Psalms embedded into {html_path}")

    print(f"[Success] Catalog synced: {len(hymns)} hymns written to {json_path}")
    return hymns

if __name__ == '__main__':
    sync_catalog()
