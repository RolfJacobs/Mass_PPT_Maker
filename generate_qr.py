#!/usr/bin/env python3
"""
QR Code Generator for Mass Planner Web App
Generates a printable, high-resolution QR code image for church noticeboards or choir phones.
"""

import sys
import argparse
import socket
import qrcode
from PIL import Image, ImageDraw, ImageFont

def get_local_ip():
    """Attempts to find the local network IP address."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # doesn't even have to be reachable
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP

def generate_qr(url, output_path="qr_code.png", label="Scan for Sunday Mass Hymn Setup"):
    print(f"Generating QR Code for URL: {url}")
    
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=3,
    )
    qr.add_data(url)
    qr.make(fit=True)

    img = qr.make_image(fill_color="#0f172a", back_color="white").convert('RGB')
    
    # Add optional top/bottom margin with text label if Pillow can render it
    w, h = img.size
    padding_top = 40
    padding_bottom = 50
    total_h = h + padding_top + padding_bottom
    
    canvas = Image.new('RGB', (w, total_h), 'white')
    canvas.paste(img, (0, padding_top))
    
    draw = ImageDraw.Draw(canvas)
    
    # Try using default font
    try:
        font = ImageFont.load_default()
        # Title at top
        draw.text((w // 2, 18), label, fill="#1e293b", font=font, anchor="mm")
        # URL at bottom
        draw.text((w // 2, total_h - 22), url, fill="#64748b", font=font, anchor="mm")
    except Exception:
        pass
    
    canvas.save(output_path)
    print(f"[Success] Saved QR code image to: {output_path}")

def main():
    parser = argparse.ArgumentParser(description="Generate QR code for Mass Planner Web App.")
    parser.add_argument("--url", default=None, help="The URL to encode in the QR code (e.g. GitHub Pages URL).")
    parser.add_argument("--output", default="qr_code.png", help="Output PNG path (default: qr_code.png).")
    parser.add_argument("--local", action="store_true", help="Use local network IP with port 8000 for local testing.")
    
    args = parser.parse_args()
    
    if args.local:
        local_ip = get_local_ip()
        url = f"http://{local_ip}:8000/webapp/"
    elif args.url:
        url = args.url
    else:
        # Default placeholder instruction
        url = "https://your-github-username.github.io/Mass_PPT_Maker/webapp/"
        print("[Note] No URL provided. Using placeholder URL. Pass --url <your-url> to customize.")
    
    generate_qr(url, args.output)

if __name__ == '__main__':
    main()
