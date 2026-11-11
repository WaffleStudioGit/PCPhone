#!/usr/bin/env python3
"""Dark Gruvbox Arch rice slideshow — Framework Laptop 13 Pro edition."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

BG = RGBColor(0x1D, 0x20, 0x21)
BG_SOFT = RGBColor(0x28, 0x28, 0x28)
BG_PANEL = RGBColor(0x3C, 0x38, 0x36)
FG = RGBColor(0xEB, 0xDB, 0xB2)
FG_DIM = RGBColor(0xA8, 0x99, 0x84)
ORANGE = RGBColor(0xFE, 0x80, 0x19)
YELLOW = RGBColor(0xFA, 0xBD, 0x2F)
GREEN = RGBColor(0xB8, 0xBB, 0x26)
AQUA = RGBColor(0x8E, 0xC0, 0x7C)
BLUE = RGBColor(0x83, 0xA5, 0x98)
RED = RGBColor(0xFB, 0x49, 0x34)
PURPLE = RGBColor(0xD3, 0x86, 0x9B)
GRAY = RGBColor(0x92, 0x83, 0x74)

ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
OUT = ROOT / "How-to-Rice-Arch.pptx"
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def set_run(run, text, size=20, color=FG, bold=False, font="Calibri"):
    run.text = text
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = rPr.makeelement(qn(f"a:{tag}"), {})
            rPr.append(el)
        el.set("typeface", font)


def add_textbox(slide, left, top, width, height, text, size=20, color=FG, bold=False, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(p.add_run(), text, size=size, color=color, bold=bold, font=font)
    return box


def fill_slide(slide, color=BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def accent_bar(slide, color=ORANGE, left=0, top=0, width=None, height=Inches(0.12)):
    width = width or SLIDE_W
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def panel(slide, left, top, width, height, color=BG_PANEL):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    try:
        shape.adjustments[0] = 0.08
    except Exception:
        pass
    return shape


def bullet_block(slide, left, top, width, height, items, size=18, color=FG, bullet_color=ORANGE):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(8)
        p.space_after = Pt(4)
        set_run(p.add_run(), "▸  ", size=size, color=bullet_color, bold=True)
        set_run(p.add_run(), item, size=size, color=color)
    return box


def code_block(slide, left, top, width, height, lines, size=14):
    panel(slide, left, top, width, height, BG_SOFT)
    box = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.4), height - Inches(0.3))
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(2)
        p.space_after = Pt(2)
        set_run(p.add_run(), line, size=size, color=AQUA, font="Consolas")


def footer(slide, page, total, label="Rice Arch on Framework 13 Pro"):
    add_textbox(slide, Inches(0.5), Inches(7.05), Inches(10), Inches(0.3), f"{label}  ·  dark gruvbox", size=11, color=GRAY)
    add_textbox(slide, Inches(11.5), Inches(7.05), Inches(1.4), Inches(0.3), f"{page} / {total}", size=11, color=GRAY, align=PP_ALIGN.RIGHT)


def section_chip(slide, left, top, text, color=ORANGE):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.08), Inches(0.32))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()
    add_textbox(slide, left + Inches(0.18), top - Inches(0.02), Inches(8), Inches(0.35), text, size=14, color=color, bold=True)


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    def title_slide(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE, height=Inches(0.18))
        accent_bar(slide, YELLOW, top=Inches(0.18), height=Inches(0.06))
        add_textbox(slide, Inches(0.9), Inches(1.55), Inches(11.5), Inches(0.4), "FRAMEWORK LAPTOP 13 PRO  ·  ARCH LINUX", size=16, color=ORANGE, bold=True)
        add_textbox(slide, Inches(0.9), Inches(2.1), Inches(11.5), Inches(1.1), "How to Rice Arch", size=52, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.9),
            Inches(3.3),
            Inches(11),
            Inches(0.9),
            "Dark Gruvbox Hyprland on the Framework 13 Pro — HiDPI, brightness,\nfingerprint, and battery tuned for this machine.\nTwo levels: Level 1 prebuilt arch-hyprland  ·  Level 2 DIY with AI",
            size=18,
            color=FG_DIM,
        )
        colors = [BG_SOFT, BG_PANEL, RED, GREEN, YELLOW, BLUE, PURPLE, AQUA, ORANGE, FG]
        x = Inches(0.9)
        for c in colors:
            s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(5.35), Inches(0.55), Inches(0.55))
            s.fill.solid()
            s.fill.fore_color.rgb = c
            s.line.fill.background()
            x += Inches(0.65)
        add_textbox(slide, Inches(0.9), Inches(6.1), Inches(11), Inches(0.35), "2.8K panel · Wayland · Google Slides compatible PPTX", size=14, color=GRAY)
        footer(slide, n, total)

    def why_fw(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, BLUE)
        section_chip(slide, Inches(0.7), Inches(0.55), "THIS MACHINE", BLUE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.6), "Why Framework 13 Pro needs its own rice notes", size=30, color=FG, bold=True)
        cards = [
            ("2.8K / HiDPI", "2880×1920 (and older 2256×1504) need real Hyprland monitor scale — not “looks fine at 100%”.", ORANGE),
            ("Laptop power", "brightnessctl, idle lock, and ALS quirks matter more than on a desktop rice.", YELLOW),
            ("Repairable Linux", "fwupd BIOS, fingerprint, Wi-Fi, and expansion cards — Arch Wiki + Framework docs first.", AQUA),
        ]
        x = Inches(0.7)
        for title, body, accent in cards:
            panel(slide, x, Inches(2.0), Inches(3.8), Inches(4.2), BG_PANEL)
            accent_bar(slide, accent, left=x, top=Inches(2.0), width=Inches(3.8), height=Inches(0.1))
            add_textbox(slide, x + Inches(0.25), Inches(2.35), Inches(3.3), Inches(0.6), title, size=20, color=accent, bold=True)
            add_textbox(slide, x + Inches(0.25), Inches(3.15), Inches(3.3), Inches(2.6), body, size=16, color=FG)
            x += Inches(4.05)
        footer(slide, n, total)

    def two_levels(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, PURPLE)
        section_chip(slide, Inches(0.7), Inches(0.55), "CHOOSE YOUR PATH", PURPLE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.6), "Two levels on the same Framework", size=32, color=FG, bold=True)
        panel(slide, Inches(0.7), Inches(1.9), Inches(5.7), Inches(4.5), BG_PANEL)
        accent_bar(slide, ORANGE, left=Inches(0.7), top=Inches(1.9), width=Inches(5.7), height=Inches(0.12))
        add_textbox(slide, Inches(0.95), Inches(2.25), Inches(5.2), Inches(0.35), "LEVEL 1", size=14, color=ORANGE, bold=True)
        add_textbox(slide, Inches(0.95), Inches(2.65), Inches(5.2), Inches(0.5), "Prebuilt: arch-hyprland", size=22, color=FG, bold=True)
        bullet_block(
            slide,
            Inches(0.95),
            Inches(3.3),
            Inches(5.2),
            Inches(2.8),
            [
                "Clone binnewbs/arch-hyprland",
                "Matugen + swww theme engine ready",
                "Then patch monitor scale for 2.8K",
                "Fix home paths + Framework scripts",
            ],
            size=15,
            bullet_color=ORANGE,
        )
        panel(slide, Inches(6.85), Inches(1.9), Inches(5.7), Inches(4.5), BG_PANEL)
        accent_bar(slide, AQUA, left=Inches(6.85), top=Inches(1.9), width=Inches(5.7), height=Inches(0.12))
        add_textbox(slide, Inches(7.1), Inches(2.25), Inches(5.2), Inches(0.35), "LEVEL 2", size=14, color=AQUA, bold=True)
        add_textbox(slide, Inches(7.1), Inches(2.65), Inches(5.2), Inches(0.5), "Make your own with AI", size=22, color=FG, bold=True)
        bullet_block(
            slide,
            Inches(7.1),
            Inches(3.3),
            Inches(5.2),
            Inches(2.8),
            [
                "Tune scale, brightness, fingerprint yourself",
                "Cursor helps write Hyprland / Waybar diffs",
                "Learn the Framework ArchWiki quirks",
                "Ship your own FW13 Pro dotfiles",
            ],
            size=15,
            bullet_color=AQUA,
        )
        footer(slide, n, total)

    def prereqs(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, GREEN)
        section_chip(slide, Inches(0.7), Inches(0.55), "BEFORE YOU START", GREEN)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Framework 13 Pro prerequisites", size=30, color=FG, bold=True)
        panel(slide, Inches(0.7), Inches(1.8), Inches(5.7), Inches(4.5), BG_PANEL)
        panel(slide, Inches(6.85), Inches(1.8), Inches(5.7), Inches(4.5), BG_PANEL)
        add_textbox(slide, Inches(0.95), Inches(2.05), Inches(5), Inches(0.4), "Hardware / firmware", size=18, color=YELLOW, bold=True)
        add_textbox(slide, Inches(7.1), Inches(2.05), Inches(5), Inches(0.4), "Software baseline", size=18, color=YELLOW, bold=True)
        bullet_block(
            slide,
            Inches(0.95),
            Inches(2.6),
            Inches(5.2),
            Inches(3.4),
            [
                "Framework Laptop 13 Pro (Intel Ultra 3 or AMD AI 300)",
                "BIOS current via fwupd / LVFS",
                "Know your panel: 2.8K or 2256×1504",
                "Wi-Fi working in installer / live ISO",
                "Backup of any existing ~/.config",
            ],
            size=15,
            bullet_color=GREEN,
        )
        bullet_block(
            slide,
            Inches(7.1),
            Inches(2.6),
            Inches(5.2),
            Inches(3.4),
            [
                "Arch (or Arch-based) with a sudo user",
                "Kernel recent (6.19+ / preferably latest)",
                "yay or paru for AUR",
                "JetBrains Mono Nerd Font",
                "Read ArchWiki: Framework Laptop 13",
            ],
            size=15,
            bullet_color=GREEN,
        )
        footer(slide, n, total)

    def hidpi(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "FRAMEWORK MUST-DO", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "HiDPI + refresh on Hyprland", size=30, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.7),
            Inches(1.6),
            Inches(12),
            Inches(0.5),
            "Put this in ~/.config/hypr/hyprland.conf (adjust connector name with hyprctl monitors).",
            size=15,
            color=FG_DIM,
        )
        code_block(
            slide,
            Inches(0.7),
            Inches(2.2),
            Inches(12),
            Inches(3.5),
            [
                "# 2.8K Framework 13 Pro panel — start at scale 2, or try 1.6–1.75 fractional",
                "monitor = eDP-1, 2880x1920@120, 0x0, 1.666667",
                "",
                "# Older 2256x1504 panel — ArchWiki suggests ~1.5",
                "# monitor = eDP-1, 2256x1504@60, 0x0, 1.5",
                "",
                "# Optional: sharper XWayland apps",
                "xwayland { force_zero_scaling = true }",
                "env = GDK_SCALE,2",
            ],
            size=15,
        )
        add_textbox(
            slide,
            Inches(0.7),
            Inches(5.95),
            Inches(12),
            Inches(0.5),
            "Tip: if fonts look huge/tiny after Level 1 (prebuilt rice), this monitor line is usually the first fix.",
            size=15,
            color=YELLOW,
        )
        footer(slide, n, total)

    def brightness_fp(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, YELLOW)
        section_chip(slide, Inches(0.7), Inches(0.55), "FRAMEWORK MUST-DO", YELLOW)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Brightness keys + fingerprint", size=30, color=FG, bold=True)
        panel(slide, Inches(0.7), Inches(1.8), Inches(5.7), Inches(4.5), BG_PANEL)
        panel(slide, Inches(6.85), Inches(1.8), Inches(5.7), Inches(4.5), BG_PANEL)
        add_textbox(slide, Inches(0.95), Inches(2.1), Inches(5.2), Inches(0.4), "Fn brightness not working?", size=17, color=YELLOW, bold=True)
        bullet_block(
            slide,
            Inches(0.95),
            Inches(2.65),
            Inches(5.2),
            Inches(1.4),
            [
                "ALS can steal brightness keys",
                "Blacklist hid_sensor_hub, reboot",
                "Bind brightnessctl in Hyprland",
            ],
            size=14,
            bullet_color=YELLOW,
        )
        code_block(
            slide,
            Inches(0.95),
            Inches(4.3),
            Inches(5.2),
            Inches(1.7),
            [
                "# /etc/modprobe.d/framework-als.conf",
                "blacklist hid_sensor_hub",
            ],
            size=13,
        )
        add_textbox(slide, Inches(7.1), Inches(2.1), Inches(5.2), Inches(0.4), "Fingerprint login", size=17, color=AQUA, bold=True)
        code_block(
            slide,
            Inches(7.1),
            Inches(2.65),
            Inches(5.2),
            Inches(3.3),
            [
                "sudo pacman -S fprintd",
                "fprintd-enroll",
                "fprintd-verify",
                "",
                "# Unenroll Windows prints first",
                "# if fprintd hangs on enroll",
            ],
            size=13,
        )
        footer(slide, n, total)

    def stack(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, BLUE)
        section_chip(slide, Inches(0.7), Inches(0.55), "THE STACK", BLUE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Hyprland stack that fits the 13 Pro", size=30, color=FG, bold=True)
        rows = [
            ("Hyprland", "Compositor — scale, gaps, blur, laptop keybinds", ORANGE),
            ("Waybar", "Battery %, brightness, network, workspaces", YELLOW),
            ("Kitty", "GPU terminal with Gruvbox / Matugen colors", AQUA),
            ("Rofi", "Launcher + wallpaper picker", BLUE),
            ("SwayNC", "Notifications / quick settings", PURPLE),
            ("swww + Matugen", "Wallpaper + auto theme (Level 1 prebuilt secret sauce)", GREEN),
            ("hyprlock / hypridle / brightnessctl", "Lock, idle, and Fn brightness on battery", RED),
        ]
        y = Inches(1.7)
        for name, desc, accent in rows:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), y, Inches(0.12), Inches(0.55))
            bar.fill.solid()
            bar.fill.fore_color.rgb = accent
            bar.line.fill.background()
            add_textbox(slide, Inches(1.05), y + Inches(0.05), Inches(3.6), Inches(0.45), name, size=15, color=accent, bold=True)
            add_textbox(slide, Inches(4.8), y + Inches(0.05), Inches(7.7), Inches(0.45), desc, size=15, color=FG)
            y += Inches(0.68)
        footer(slide, n, total)

    def palette(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "VISUAL LANGUAGE", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Dark Gruvbox palette", size=30, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.7),
            Inches(1.55),
            Inches(12),
            Inches(0.4),
            "Level 1: Matugen derives colors from wallpaper. Level 2 DIY: lock these Gruvbox hex values by hand.",
            size=15,
            color=FG_DIM,
        )
        swatches = [
            ("bg0_h", "#1d2021", BG),
            ("bg", "#282828", BG_SOFT),
            ("bg1", "#3c3836", BG_PANEL),
            ("fg", "#ebdbb2", FG),
            ("gray", "#928374", GRAY),
            ("red", "#fb4934", RED),
            ("green", "#b8bb26", GREEN),
            ("yellow", "#fabd2f", YELLOW),
            ("blue", "#83a598", BLUE),
            ("purple", "#d3869b", PURPLE),
            ("aqua", "#8ec07c", AQUA),
            ("orange", "#fe8019", ORANGE),
        ]
        x0, y0 = Inches(0.7), Inches(2.2)
        for i, (name, hexv, col) in enumerate(swatches):
            col_i, row_i = i % 6, i // 6
            x = x0 + Inches(col_i * 2.05)
            y = y0 + Inches(row_i * 2.0)
            panel(slide, x, y, Inches(1.9), Inches(1.7), BG_PANEL)
            s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.2), y + Inches(0.25), Inches(1.5), Inches(0.7))
            s.fill.solid()
            s.fill.fore_color.rgb = col
            s.line.fill.background()
            add_textbox(slide, x + Inches(0.15), y + Inches(1.05), Inches(1.6), Inches(0.3), name, size=14, color=FG, bold=True, align=PP_ALIGN.CENTER)
            add_textbox(slide, x + Inches(0.15), y + Inches(1.3), Inches(1.6), Inches(0.3), hexv, size=12, color=FG_DIM, align=PP_ALIGN.CENTER)
        footer(slide, n, total)

    def level1_divider(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE, height=Inches(0.2))
        add_textbox(slide, Inches(0.9), Inches(2.2), Inches(11), Inches(0.4), "LEVEL 1  ·  FRAMEWORK 13 PRO", size=16, color=ORANGE, bold=True)
        add_textbox(slide, Inches(0.9), Inches(2.75), Inches(11.5), Inches(1.0), "Prebuilt: binnewbs/arch-hyprland", size=36, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.9),
            Inches(4.0),
            Inches(11),
            Inches(0.9),
            "Start here. Drop in the Matugen rice, then patch it for Framework HiDPI,\nhome paths, brightness, and fingerprint.",
            size=18,
            color=FG_DIM,
        )
        footer(slide, n, total)

    def level1_look(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 1 · LOOK", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.5), "What you’re installing (then adapting)", size=28, color=FG, bold=True)
        img1, img2 = ASSETS / "rice1.png", ASSETS / "rice2.png"
        if img1.exists():
            slide.shapes.add_picture(str(img1), Inches(0.7), Inches(1.65), width=Inches(5.9))
        if img2.exists():
            slide.shapes.add_picture(str(img2), Inches(6.8), Inches(1.65), width=Inches(5.9))
        add_textbox(
            slide,
            Inches(0.7),
            Inches(6.3),
            Inches(12),
            Inches(0.35),
            "github.com/binnewbs/arch-hyprland  ·  on FW13 Pro, fix eDP scale before judging fonts",
            size=13,
            color=GRAY,
        )
        footer(slide, n, total)

    def level1_install(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 1 · INSTALL", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.5), "Clone, then Framework-patch", size=28, color=FG, bold=True)
        code_block(
            slide,
            Inches(0.7),
            Inches(1.6),
            Inches(12),
            Inches(3.0),
            [
                "cp -a ~/.config ~/.config.bak.$(date +%F)",
                "git clone https://github.com/binnewbs/arch-hyprland.git ~/arch-hyprland",
                "cd ~/arch-hyprland && cp -r .config/* ~/.config/",
                "mkdir -p ~/Pictures && cp -r wallpapers ~/Pictures/wallpapers",
            ],
            size=14,
        )
        bullet_block(
            slide,
            Inches(0.7),
            Inches(4.85),
            Inches(12),
            Inches(1.7),
            [
                "Replace /home/beentang/... paths with your username",
                "Set monitor = eDP-1,2880x1920@120,0x0,<scale> for the 13 Pro panel",
                "Wire brightnessctl into existing keybinds; enroll fprintd",
                "Run wppicker / matugen image after packages are installed",
            ],
            size=15,
            bullet_color=ORANGE,
        )
        footer(slide, n, total)

    def level1_matugen(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 1 · SECRET SAUCE", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.5), "Matugen + swww on a laptop", size=28, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.7),
            Inches(1.55),
            Inches(12),
            Inches(0.45),
            "Colors follow wallpaper — great for Gruvbox-ish vibes without hand-editing every CSS file.",
            size=15,
            color=FG_DIM,
        )
        panels = [
            ("1. Pick wallpaper", "wppicker.sh → Rofi from ~/Pictures/wallpapers", ORANGE),
            ("2. Matugen image", "Writes Waybar, Kitty, Hypr, Rofi, GTK…", YELLOW),
            ("3. Live reload", "Hooks restart bar / terminal / hyprctl", GREEN),
            ("4. FW13 check", "Re-check scale + battery modules still show", BLUE),
        ]
        x = Inches(0.7)
        for title, body, accent in panels:
            panel(slide, x, Inches(2.3), Inches(2.95), Inches(3.6), BG_PANEL)
            accent_bar(slide, accent, left=x, top=Inches(2.3), width=Inches(2.95), height=Inches(0.1))
            add_textbox(slide, x + Inches(0.2), Inches(2.65), Inches(2.55), Inches(0.7), title, size=16, color=accent, bold=True)
            add_textbox(slide, x + Inches(0.2), Inches(3.5), Inches(2.55), Inches(2.0), body, size=14, color=FG)
            x += Inches(3.15)
        footer(slide, n, total)

    def level2_divider(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA, height=Inches(0.2))
        add_textbox(slide, Inches(0.9), Inches(2.2), Inches(11), Inches(0.4), "LEVEL 2  ·  FRAMEWORK 13 PRO", size=16, color=AQUA, bold=True)
        add_textbox(slide, Inches(0.9), Inches(2.75), Inches(11.5), Inches(1.0), "Make your own rice with AI", size=40, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.9),
            Inches(4.0),
            Inches(11),
            Inches(0.9),
            "Go deeper. You own every config. Cursor writes the diffs while you dial in\nscale, battery, and Gruvbox for this laptop.",
            size=18,
            color=FG_DIM,
        )
        footer(slide, n, total)

    def level2_steps(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 2 · WORKFLOW", AQUA)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Build loop on the 13 Pro", size=30, color=FG, bold=True)
        steps = [
            ("01", "Install Arch + firmware", "fwupdmgr get-updates; confirm Wi-Fi + graphics"),
            ("02", "Hyprland bare metal", "Terminal + keybinds first — then scale the eDP panel"),
            ("03", "Laptop controls", "brightnessctl, hypridle/hyprlock, battery in Waybar"),
            ("04", "Lock Gruvbox", "Kitty + Waybar + borders on the same palette"),
            ("05", "Iterate with AI", "Paste hyprctl monitors, errors, screenshots to Cursor"),
            ("06", "Commit FW13 dots", "Keep hypr / waybar / kitty in git for the next mainboard"),
        ]
        y = Inches(1.7)
        for num, title, body in steps:
            add_textbox(slide, Inches(0.7), y, Inches(0.7), Inches(0.4), num, size=18, color=AQUA, bold=True)
            add_textbox(slide, Inches(1.5), y, Inches(3.8), Inches(0.4), title, size=16, color=FG, bold=True)
            add_textbox(slide, Inches(5.5), y, Inches(7), Inches(0.4), body, size=15, color=FG_DIM)
            y += Inches(0.75)
        footer(slide, n, total)

    def level2_packages(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 2 · PACKAGES", AQUA)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Starter set for Framework + Hyprland", size=28, color=FG, bold=True)
        code_block(
            slide,
            Inches(0.7),
            Inches(1.7),
            Inches(12),
            Inches(4.7),
            [
                "sudo pacman -S hyprland waybar kitty rofi-wayland swww \\",
                "  hyprlock hypridle grim slurp wl-clipboard brightnessctl \\",
                "  playerctl pipewire wireplumber networkmanager \\",
                "  fprintd power-profiles-daemon fwupd \\",
                "  ttf-jetbrains-mono-nerd noto-fonts-emoji",
                "",
                "yay -S swaync-git wlogout matugen-bin   # optional polish",
                "",
                "# Then: set eDP-1 scale, bind brightnessctl, enroll fingerprint,",
                "# and only after that chase the Gruvbox look.",
            ],
            size=15,
        )
        footer(slide, n, total)

    def level2_prompts(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 2 · AI PROMPTS", AQUA)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Prompts that mention the Framework", size=28, color=FG, bold=True)
        prompts = [
            "“Framework 13 Pro 2.8K — write a Hyprland monitor line for 2880x1920@120 with ~1.67 scale.”",
            "“Bind XF86MonBrightnessUp/Down to brightnessctl on Hyprland; Waybar show battery + backlight.”",
            "“Gruvbox Dark for kitty + waybar; laptop rice, not desktop — keep contrast readable outdoors.”",
            "“Add hypridle + hyprlock using my current wallpaper symlink; lock after 5 min on battery.”",
            "“Here’s hyprctl monitors and my config — fix blurry XWayland apps after fractional scaling.”",
        ]
        y = Inches(1.7)
        for i, text in enumerate(prompts, 1):
            panel(slide, Inches(0.7), y, Inches(12), Inches(0.85), BG_PANEL)
            add_textbox(slide, Inches(0.95), y + Inches(0.22), Inches(0.5), Inches(0.4), f"{i}.", size=16, color=AQUA, bold=True)
            add_textbox(slide, Inches(1.4), y + Inches(0.18), Inches(11), Inches(0.55), text, size=14, color=FG)
            y += Inches(0.95)
        footer(slide, n, total)

    def compare(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, PURPLE)
        section_chip(slide, Inches(0.7), Inches(0.55), "PICK A LANE", PURPLE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.5), "Level 1 vs Level 2 on Framework 13 Pro", size=28, color=FG, bold=True)
        headers = ["", "Level 1 · arch-hyprland", "Level 2 · DIY + AI"]
        rows = [
            ("Speed to pretty UI", "Fast — then patch scale", "Slower"),
            ("FW13 Pro quirks", "Must add after clone", "You wire them first"),
            ("Learning", "Medium", "Highest"),
            ("Best when…", "You want this look tonight", "You want a signature FW rice"),
        ]
        add_textbox(slide, Inches(0.8), Inches(1.7), Inches(3.5), Inches(0.4), headers[0], size=14, color=GRAY, bold=True)
        add_textbox(slide, Inches(4.4), Inches(1.7), Inches(4), Inches(0.4), headers[1], size=16, color=ORANGE, bold=True)
        add_textbox(slide, Inches(8.6), Inches(1.7), Inches(4), Inches(0.4), headers[2], size=16, color=AQUA, bold=True)
        y = Inches(2.3)
        for i, (a, b, c) in enumerate(rows):
            panel(slide, Inches(0.7), y, Inches(12), Inches(0.85), BG_PANEL if i % 2 == 0 else BG_SOFT)
            add_textbox(slide, Inches(0.9), y + Inches(0.2), Inches(3.3), Inches(0.45), a, size=15, color=FG_DIM)
            add_textbox(slide, Inches(4.4), y + Inches(0.2), Inches(4), Inches(0.45), b, size=15, color=FG)
            add_textbox(slide, Inches(8.6), y + Inches(0.2), Inches(4), Inches(0.45), c, size=15, color=FG)
            y += Inches(0.95)
        footer(slide, n, total)

    def checklist(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, GREEN)
        section_chip(slide, Inches(0.7), Inches(0.55), "SHIP IT", GREEN)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.5), "Framework 13 Pro rice checklist", size=28, color=FG, bold=True)
        items = [
            ("Display", "eDP scale set; external monitor rule ready if you dock"),
            ("Input", "Touchpad tap-to-click; natural scroll preference"),
            ("Power", "power-profiles-daemon or TLP; hypridle on battery"),
            ("Brightness", "Fn keys via brightnessctl (ALS blacklisted if needed)"),
            ("Fingerprint", "fprintd enrolled; works at login / sudo if you want"),
            ("Firmware", "fwupdmgr refresh && fwupdmgr update"),
            ("Theme", "Gruvbox locked OR Matugen regenerates cleanly"),
            ("Backup", "Dotfiles in git — Framework boards get upgraded"),
        ]
        y = Inches(1.65)
        for title, body in items:
            add_textbox(slide, Inches(0.7), y, Inches(2.4), Inches(0.45), title, size=15, color=GREEN, bold=True)
            add_textbox(slide, Inches(3.3), y, Inches(9.2), Inches(0.45), body, size=15, color=FG)
            y += Inches(0.58)
        footer(slide, n, total)

    def resources(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, BLUE)
        section_chip(slide, Inches(0.7), Inches(0.55), "LINKS", BLUE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.5), "Resources", size=28, color=FG, bold=True)
        links = [
            ("Framework 13 Pro", "https://frame.work/laptop13pro"),
            ("ArchWiki · Framework 13", "https://wiki.archlinux.org/title/Framework_Laptop_13"),
            ("Level 1 rice", "https://github.com/binnewbs/arch-hyprland"),
            ("Hyprland wiki", "https://wiki.hyprland.org"),
            ("Gruvbox", "https://github.com/morhetz/gruvbox"),
            ("Matugen", "https://github.com/InioX/matugen"),
        ]
        y = Inches(1.7)
        for name, url in links:
            panel(slide, Inches(0.7), y, Inches(12), Inches(0.7), BG_PANEL)
            add_textbox(slide, Inches(0.95), y + Inches(0.15), Inches(4), Inches(0.4), name, size=16, color=BLUE, bold=True)
            add_textbox(slide, Inches(5.1), y + Inches(0.15), Inches(7.3), Inches(0.4), url, size=14, color=FG_DIM)
            y += Inches(0.8)
        footer(slide, n, total)

    def end_slide(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE, height=Inches(0.18))
        accent_bar(slide, AQUA, top=Inches(0.18), height=Inches(0.06))
        add_textbox(slide, Inches(0.9), Inches(2.2), Inches(11.5), Inches(0.4), "FRAMEWORK 13 PRO", size=16, color=ORANGE, bold=True)
        add_textbox(slide, Inches(0.9), Inches(2.7), Inches(11.5), Inches(1.0), "Pick a level. Rice the laptop.", size=38, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.9),
            Inches(3.9),
            Inches(11),
            Inches(1.1),
            "Level 1 — start from arch-hyprland, then patch for this Framework.\nLevel 2 — design with AI until scale, power, and Gruvbox feel native.",
            size=17,
            color=FG_DIM,
        )
        add_textbox(
            slide,
            Inches(0.9),
            Inches(5.4),
            Inches(11),
            Inches(0.5),
            "Open this .pptx in Google Slides: Drive → Open with Google Slides",
            size=15,
            color=GRAY,
        )
        footer(slide, n, total)

    builders = [
        title_slide,
        why_fw,
        two_levels,
        prereqs,
        hidpi,
        brightness_fp,
        stack,
        palette,
        level1_divider,
        level1_look,
        level1_install,
        level1_matugen,
        level2_divider,
        level2_steps,
        level2_packages,
        level2_prompts,
        compare,
        checklist,
        resources,
        end_slide,
    ]

    total = len(builders)
    for i, fn in enumerate(builders, 1):
        slide = prs.slides.add_slide(blank)
        fn(slide, i, total)

    prs.save(OUT)
    print(f"Wrote {OUT} ({total} slides)")


if __name__ == "__main__":
    build()
