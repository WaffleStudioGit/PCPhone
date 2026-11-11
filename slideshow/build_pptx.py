#!/usr/bin/env python3
"""Build a dark Gruvbox Arch rice slideshow (Google Slides–importable PPTX)."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

# --- Gruvbox Dark ---
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
    # Help Google Slides map fonts more cleanly
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


def add_para(tf, text, size=18, color=FG, bold=False, space_before=6, space_after=4, font="Calibri"):
    p = tf.add_paragraph()
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    set_run(p.add_run(), text, size=size, color=color, bold=bold, font=font)
    return p


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
    shape = panel(slide, left, top, width, height, BG_SOFT)
    box = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.4), height - Inches(0.3))
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(2)
        p.space_after = Pt(2)
        set_run(p.add_run(), line, size=size, color=AQUA, font="Consolas")
    return shape


def footer(slide, page, total, label="How to Rice Arch"):
    add_textbox(
        slide,
        Inches(0.5),
        Inches(7.05),
        Inches(10),
        Inches(0.3),
        f"{label}  ·  dark gruvbox",
        size=11,
        color=GRAY,
    )
    add_textbox(
        slide,
        Inches(11.5),
        Inches(7.05),
        Inches(1.4),
        Inches(0.3),
        f"{page} / {total}",
        size=11,
        color=GRAY,
        align=PP_ALIGN.RIGHT,
    )


def section_chip(slide, left, top, text, color=ORANGE):
    # small label strip
    w = Inches(0.08)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, w, Inches(0.32))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()
    add_textbox(slide, left + Inches(0.18), top - Inches(0.02), Inches(6), Inches(0.35), text, size=14, color=color, bold=True)


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    slides_meta = []  # collect for numbering at end; we'll rebuild with known total

    # We'll construct slides in order into a list of builders
    builders = []

    def title_slide(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE, height=Inches(0.18))
        accent_bar(slide, YELLOW, top=Inches(0.18), height=Inches(0.06))
        add_textbox(slide, Inches(0.9), Inches(1.8), Inches(11.5), Inches(0.5), "ARCH LINUX", size=18, color=ORANGE, bold=True)
        add_textbox(slide, Inches(0.9), Inches(2.3), Inches(11.5), Inches(1.2), "How to Rice Arch", size=54, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.9),
            Inches(3.5),
            Inches(11),
            Inches(0.8),
            "A dark Gruvbox walkthrough with two levels:\nbuild your own with AI  ·  or install binnewbs/arch-hyprland",
            size=20,
            color=FG_DIM,
        )
        # palette swatches
        colors = [BG_SOFT, BG_PANEL, RED, GREEN, YELLOW, BLUE, PURPLE, AQUA, ORANGE, FG]
        x = Inches(0.9)
        for c in colors:
            s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(5.4), Inches(0.55), Inches(0.55))
            s.fill.solid()
            s.fill.fore_color.rgb = c
            s.line.fill.background()
            x += Inches(0.65)
        add_textbox(slide, Inches(0.9), Inches(6.15), Inches(10), Inches(0.35), "Theme: Gruvbox Dark  ·  Target: Hyprland on Arch  ·  Google Slides compatible", size=14, color=GRAY)
        footer(slide, n, total)

    def what_is_rice(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, BLUE)
        section_chip(slide, Inches(0.7), Inches(0.55), "BASICS")
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(11), Inches(0.7), "What “ricing” means", size=36, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.7),
            Inches(1.7),
            Inches(11.5),
            Inches(0.7),
            "Ricing is customizing your desktop until it looks intentional — colors, fonts, bar, terminal, wallpaper, and animations all speaking the same language.",
            size=18,
            color=FG_DIM,
        )
        cards = [
            ("Compositor", "Hyprland manages windows, gaps, blur, and keybinds.", ORANGE),
            ("Bar + launcher", "Waybar + Rofi give you status and app launch.", YELLOW),
            ("Terminal + theme", "Kitty + Gruvbox/Matugen keep everything cohesive.", AQUA),
        ]
        x = Inches(0.7)
        for title, body, accent in cards:
            panel(slide, x, Inches(2.8), Inches(3.8), Inches(3.2), BG_PANEL)
            accent_bar(slide, accent, left=x, top=Inches(2.8), width=Inches(3.8), height=Inches(0.1))
            add_textbox(slide, x + Inches(0.25), Inches(3.15), Inches(3.3), Inches(0.5), title, size=20, color=accent, bold=True)
            add_textbox(slide, x + Inches(0.25), Inches(3.75), Inches(3.3), Inches(1.8), body, size=16, color=FG)
            x += Inches(4.05)
        footer(slide, n, total)

    def two_levels(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, PURPLE)
        section_chip(slide, Inches(0.7), Inches(0.55), "CHOOSE YOUR PATH", PURPLE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.7), "Two levels, one slideshow", size=36, color=FG, bold=True)

        # Level 1
        panel(slide, Inches(0.7), Inches(2.0), Inches(5.7), Inches(4.4), BG_PANEL)
        accent_bar(slide, ORANGE, left=Inches(0.7), top=Inches(2.0), width=Inches(5.7), height=Inches(0.12))
        add_textbox(slide, Inches(0.95), Inches(2.35), Inches(5.2), Inches(0.4), "LEVEL 1", size=14, color=ORANGE, bold=True)
        add_textbox(slide, Inches(0.95), Inches(2.75), Inches(5.2), Inches(0.5), "Make your own rice with AI", size=24, color=FG, bold=True)
        bullet_block(
            slide,
            Inches(0.95),
            Inches(3.4),
            Inches(5.2),
            Inches(2.7),
            [
                "You design the look piece by piece",
                "Cursor / chat helps write configs",
                "Learn Hyprland, Waybar, Kitty deeply",
                "Ends as your own dotfiles repo",
            ],
            size=16,
            bullet_color=ORANGE,
        )

        # Level 2
        panel(slide, Inches(6.85), Inches(2.0), Inches(5.7), Inches(4.4), BG_PANEL)
        accent_bar(slide, AQUA, left=Inches(6.85), top=Inches(2.0), width=Inches(5.7), height=Inches(0.12))
        add_textbox(slide, Inches(7.1), Inches(2.35), Inches(5.2), Inches(0.4), "LEVEL 2", size=14, color=AQUA, bold=True)
        add_textbox(slide, Inches(7.1), Inches(2.75), Inches(5.2), Inches(0.5), "Use arch-hyprland", size=24, color=FG, bold=True)
        bullet_block(
            slide,
            Inches(7.1),
            Inches(3.4),
            Inches(5.2),
            Inches(2.7),
            [
                "Clone binnewbs/arch-hyprland",
                "Matugen colors from wallpapers",
                "Waybar, Rofi, SwayNC, Kitty ready",
                "Tweak paths, then make it yours",
            ],
            size=16,
            bullet_color=AQUA,
        )
        footer(slide, n, total)

    def prereqs(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, GREEN)
        section_chip(slide, Inches(0.7), Inches(0.55), "BEFORE YOU START", GREEN)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.7), "Prerequisites", size=36, color=FG, bold=True)
        left_items = [
            "Arch Linux (or Arch-based) installed",
            "A normal user with sudo",
            "Working network + graphics drivers",
            "AUR helper: yay or paru",
            "Backup of any existing ~/.config",
        ]
        right_items = [
            "Know how to edit files in a terminal",
            "Comfortable rebooting into a TTY if needed",
            "Wallpapers folder ready (Pictures/wallpapers)",
            "Nerd Font installed (JetBrains Mono recommended)",
            "Time to iterate — rice is iterative",
        ]
        panel(slide, Inches(0.7), Inches(1.9), Inches(5.7), Inches(4.5), BG_PANEL)
        panel(slide, Inches(6.85), Inches(1.9), Inches(5.7), Inches(4.5), BG_PANEL)
        add_textbox(slide, Inches(0.95), Inches(2.15), Inches(5), Inches(0.4), "System", size=18, color=YELLOW, bold=True)
        add_textbox(slide, Inches(7.1), Inches(2.15), Inches(5), Inches(0.4), "Mindset", size=18, color=YELLOW, bold=True)
        bullet_block(slide, Inches(0.95), Inches(2.7), Inches(5.2), Inches(3.4), left_items, size=16, bullet_color=GREEN)
        bullet_block(slide, Inches(7.1), Inches(2.7), Inches(5.2), Inches(3.4), right_items, size=16, bullet_color=GREEN)
        footer(slide, n, total)

    def stack(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, BLUE)
        section_chip(slide, Inches(0.7), Inches(0.55), "THE STACK", BLUE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.6), "Core pieces of a Hyprland rice", size=34, color=FG, bold=True)

        rows = [
            ("Hyprland", "Wayland compositor — windows, gaps, blur, keybinds", ORANGE),
            ("Waybar", "Top/status bar — workspaces, clock, tray, modules", YELLOW),
            ("Kitty", "GPU terminal — fonts + Gruvbox / Matugen colors", AQUA),
            ("Rofi", "App launcher + wallpaper picker UI", BLUE),
            ("SwayNC", "Notification center / control panel", PURPLE),
            ("swww + Matugen", "Wallpaper daemon + auto color scheme generator", GREEN),
            ("hyprlock / hypridle / wlogout", "Lock, idle, and logout menu", RED),
        ]
        y = Inches(1.7)
        for name, desc, accent in rows:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), y, Inches(0.12), Inches(0.55))
            bar.fill.solid()
            bar.fill.fore_color.rgb = accent
            bar.line.fill.background()
            add_textbox(slide, Inches(1.05), y + Inches(0.05), Inches(3.2), Inches(0.45), name, size=16, color=accent, bold=True)
            add_textbox(slide, Inches(4.3), y + Inches(0.05), Inches(8.2), Inches(0.45), desc, size=16, color=FG)
            y += Inches(0.68)
        footer(slide, n, total)

    def palette(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "VISUAL LANGUAGE", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.6), "Dark Gruvbox palette", size=34, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.7),
            Inches(1.6),
            Inches(12),
            Inches(0.5),
            "Keep every app on these hex values (or let Matugen derive a matching set from wallpaper).",
            size=16,
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
        x0, y0 = Inches(0.7), Inches(2.3)
        for i, (name, hexv, col) in enumerate(swatches):
            col_i = i % 6
            row_i = i // 6
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
        add_textbox(slide, Inches(0.9), Inches(2.4), Inches(11), Inches(0.4), "LEVEL 1", size=18, color=ORANGE, bold=True)
        add_textbox(slide, Inches(0.9), Inches(2.9), Inches(11.5), Inches(1.0), "Make your own rice with AI", size=44, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.9),
            Inches(4.1),
            Inches(11),
            Inches(0.8),
            "You own every config. Cursor (or any AI) writes the boring parts while you decide the look.",
            size=18,
            color=FG_DIM,
        )
        footer(slide, n, total)

    def level1_steps(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 1 · WORKFLOW", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.6), "Build loop with an AI pair", size=32, color=FG, bold=True)
        steps = [
            ("01", "Install base packages", "hyprland waybar kitty rofi-wayland swww grim slurp wl-clipboard"),
            ("02", "Boot Hyprland bare", "Get a terminal + keybinds working before styling"),
            ("03", "Lock Gruvbox early", "Paste palette into kitty, waybar, hypr borders"),
            ("04", "Iterate with AI", "Paste errors + screenshots; ask for exact config diffs"),
            ("05", "Commit dotfiles", "Keep ~/.config/hypr waybar kitty in a git repo"),
        ]
        y = Inches(1.75)
        for num, title, body in steps:
            add_textbox(slide, Inches(0.7), y, Inches(0.8), Inches(0.45), num, size=20, color=ORANGE, bold=True)
            add_textbox(slide, Inches(1.6), y, Inches(4.2), Inches(0.45), title, size=18, color=FG, bold=True)
            add_textbox(slide, Inches(5.9), y, Inches(6.7), Inches(0.45), body, size=15, color=FG_DIM)
            y += Inches(0.85)
        footer(slide, n, total)

    def level1_packages(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 1 · PACKAGES", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.6), "Starter pacman + AUR set", size=32, color=FG, bold=True)
        code_block(
            slide,
            Inches(0.7),
            Inches(1.8),
            Inches(12),
            Inches(4.6),
            [
                "# Official repos",
                "sudo pacman -S hyprland waybar kitty rofi-wayland swww \\",
                "  hyprlock hypridle grim slurp wl-clipboard brightnessctl \\",
                "  playerctl pipewire wireplumber networkmanager \\",
                "  ttf-jetbrains-mono-nerd noto-fonts-emoji",
                "",
                "# AUR (yay/paru)",
                "yay -S swaync-git wlogout matugen-bin  # optional but nice",
                "",
                "# Then: copy a minimal hyprland.conf, open Kitty with Super+Q,",
                "# and only then start theming.",
            ],
            size=15,
        )
        footer(slide, n, total)

    def level1_ai_prompts(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 1 · AI PROMPTS", ORANGE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.6), "Good prompts to give Cursor", size=32, color=FG, bold=True)
        prompts = [
            "“Make my Hyprland borders and Waybar use Gruvbox Dark; here is my current style.css.”",
            "“Write a dwindle layout hyprland.conf with Super+Q kitty, Super+R rofi, gaps 5/10.”",
            "“Fix this Waybar JSONC — modules-right should show pulse, network, clock.”",
            "“Generate a Kitty Gruvbox dark theme matching #282828 / #ebdbb2.”",
            "“Add a wallpaper picker script using swww + rofi.”",
        ]
        y = Inches(1.8)
        for i, text in enumerate(prompts, 1):
            panel(slide, Inches(0.7), y, Inches(12), Inches(0.85), BG_PANEL)
            add_textbox(slide, Inches(0.95), y + Inches(0.22), Inches(0.5), Inches(0.4), f"{i}.", size=16, color=ORANGE, bold=True)
            add_textbox(slide, Inches(1.4), y + Inches(0.2), Inches(11), Inches(0.55), text, size=15, color=FG)
            y += Inches(0.95)
        footer(slide, n, total)

    def level2_divider(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA, height=Inches(0.2))
        add_textbox(slide, Inches(0.9), Inches(2.4), Inches(11), Inches(0.4), "LEVEL 2", size=18, color=AQUA, bold=True)
        add_textbox(slide, Inches(0.9), Inches(2.9), Inches(11.5), Inches(1.0), "Use binnewbs/arch-hyprland", size=40, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.9),
            Inches(4.1),
            Inches(11),
            Inches(0.9),
            "Personal Hyprland rice with Matugen-driven colors, swww wallpapers,\nWaybar layouts, Rofi, SwayNC, Kitty, Cava, and wlogout.",
            size=18,
            color=FG_DIM,
        )
        footer(slide, n, total)

    def level2_look(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 2 · LOOK", AQUA)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "What this rice looks like", size=32, color=FG, bold=True)
        img1 = ASSETS / "rice1.png"
        img2 = ASSETS / "rice2.png"
        if img1.exists():
            slide.shapes.add_picture(str(img1), Inches(0.7), Inches(1.7), width=Inches(5.9))
        if img2.exists():
            slide.shapes.add_picture(str(img2), Inches(6.8), Inches(1.7), width=Inches(5.9))
        add_textbox(
            slide,
            Inches(0.7),
            Inches(6.35),
            Inches(12),
            Inches(0.4),
            "Screenshots from github.com/binnewbs/arch-hyprland  ·  Matugen palette follows wallpaper",
            size=13,
            color=GRAY,
        )
        footer(slide, n, total)

    def level2_install(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 2 · INSTALL", AQUA)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Clone and place the configs", size=32, color=FG, bold=True)
        code_block(
            slide,
            Inches(0.7),
            Inches(1.7),
            Inches(12),
            Inches(3.3),
            [
                "# Backup first",
                "cp -a ~/.config ~/.config.bak.$(date +%F)",
                "",
                "git clone https://github.com/binnewbs/arch-hyprland.git ~/arch-hyprland",
                "cd ~/arch-hyprland",
                "",
                "# Merge carefully — don’t blind-overwrite paths you still need",
                "cp -r .config/* ~/.config/",
                "cp -r wallpapers ~/Pictures/wallpapers   # or symlink",
                "cp .zshrc ~/.zshrc                     # optional",
            ],
            size=14,
        )
        bullet_block(
            slide,
            Inches(0.7),
            Inches(5.2),
            Inches(12),
            Inches(1.4),
            [
                "Install matching packages: hyprland, waybar, kitty, rofi, swww, swaync, matugen, wlogout, cava, fastfetch",
                "Edit absolute paths (the upstream rice uses /home/beentang/... in places)",
                "Reload: hyprctl reload  ·  restart Waybar  ·  pick a wallpaper with the wppicker script",
            ],
            size=15,
            bullet_color=AQUA,
        )
        footer(slide, n, total)

    def level2_matugen(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 2 · SECRET SAUCE", AQUA)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Matugen + swww workflow", size=32, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.7),
            Inches(1.65),
            Inches(12),
            Inches(0.6),
            "This rice regenerates theme colors from the wallpaper so you rarely hand-edit every color file.",
            size=16,
            color=FG_DIM,
        )
        panels = [
            ("1. Pick wallpaper", "wppicker.sh → Rofi preview from ~/Pictures/wallpapers", ORANGE),
            ("2. Matugen image", "matugen image <file> writes templates for each app", YELLOW),
            ("3. Live reload", "hooks reload Waybar, Kitty, Hyprland, Cava", GREEN),
            ("4. Enjoy", "borders, bar, terminal, GTK, Rofi, Discord/Spotify templates", BLUE),
        ]
        x = Inches(0.7)
        for title, body, accent in panels:
            panel(slide, x, Inches(2.5), Inches(2.95), Inches(3.5), BG_PANEL)
            accent_bar(slide, accent, left=x, top=Inches(2.5), width=Inches(2.95), height=Inches(0.1))
            add_textbox(slide, x + Inches(0.2), Inches(2.85), Inches(2.55), Inches(0.8), title, size=16, color=accent, bold=True)
            add_textbox(slide, x + Inches(0.2), Inches(3.7), Inches(2.55), Inches(2.0), body, size=14, color=FG)
            x += Inches(3.15)
        footer(slide, n, total)

    def level2_layout(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, AQUA)
        section_chip(slide, Inches(0.7), Inches(0.55), "LEVEL 2 · REPO MAP", AQUA)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Where things live", size=32, color=FG, bold=True)
        code_block(
            slide,
            Inches(0.7),
            Inches(1.7),
            Inches(12),
            Inches(4.7),
            [
                ".config/hypr/           hyprland.conf + colors + scripts (wppicker, volume…)",
                ".config/hypr/configs/   keybinds, looknfeel, windowrules, animations",
                ".config/waybar/         layouts + style CSS + Matugen colors.css",
                ".config/kitty/          kitty.conf + colors.conf",
                ".config/matugen/        config.toml + templates for every themed app",
                ".config/rofi/           launcher theme + colors.rasi",
                ".config/swaync/         notification center style",
                ".config/wlogout/        logout menu",
                ".config/cava/           audio visualizer",
                "wallpapers/             starter wallpapers for Matugen",
            ],
            size=15,
        )
        footer(slide, n, total)

    def compare(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, PURPLE)
        section_chip(slide, Inches(0.7), Inches(0.55), "PICK A LANE", PURPLE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Level 1 vs Level 2", size=32, color=FG, bold=True)

        headers = ["", "Level 1 · DIY + AI", "Level 2 · arch-hyprland"]
        rows = [
            ("Speed to pretty desktop", "Slower", "Fast"),
            ("Learning depth", "Highest", "Medium — still tweak paths"),
            ("Maintenance", "You own every line", "Upstream + your forks"),
            ("Color system", "Hand Gruvbox or Matugen", "Matugen-first"),
            ("Best when…", "You want a signature look", "You want this rice now"),
        ]
        # header
        add_textbox(slide, Inches(0.8), Inches(1.8), Inches(3.5), Inches(0.4), headers[0], size=14, color=GRAY, bold=True)
        add_textbox(slide, Inches(4.4), Inches(1.8), Inches(4), Inches(0.4), headers[1], size=16, color=ORANGE, bold=True)
        add_textbox(slide, Inches(8.6), Inches(1.8), Inches(4), Inches(0.4), headers[2], size=16, color=AQUA, bold=True)
        y = Inches(2.4)
        for i, (a, b, c) in enumerate(rows):
            bgc = BG_PANEL if i % 2 == 0 else BG_SOFT
            panel(slide, Inches(0.7), y, Inches(12), Inches(0.7), bgc)
            add_textbox(slide, Inches(0.9), y + Inches(0.15), Inches(3.3), Inches(0.45), a, size=15, color=FG_DIM)
            add_textbox(slide, Inches(4.4), y + Inches(0.15), Inches(4), Inches(0.45), b, size=15, color=FG)
            add_textbox(slide, Inches(8.6), y + Inches(0.15), Inches(4), Inches(0.45), c, size=15, color=FG)
            y += Inches(0.75)
        footer(slide, n, total)

    def polish(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, YELLOW)
        section_chip(slide, Inches(0.7), Inches(0.55), "FINISHING", YELLOW)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Make it feel finished", size=32, color=FG, bold=True)
        items = [
            ("Fonts", "JetBrains Mono Nerd everywhere — bar, terminal, Rofi"),
            ("GTK + icons", "Match dark Gruvbox / Material; set cursor theme"),
            ("Blur & gaps", "rounding 10, gaps 5/10, inactive opacity ~0.8 (like this rice)"),
            ("Lock screen", "hyprlock wallpaper + Gruvbox text colors"),
            ("Dotfiles hygiene", "git init your ~/.config slices; never commit secrets"),
            ("Screenshot flex", "grim + slurp; keep a before/after pair"),
        ]
        y = Inches(1.75)
        for title, body in items:
            add_textbox(slide, Inches(0.7), y, Inches(2.8), Inches(0.5), title, size=16, color=YELLOW, bold=True)
            add_textbox(slide, Inches(3.6), y, Inches(9), Inches(0.5), body, size=16, color=FG)
            y += Inches(0.75)
        footer(slide, n, total)

    def resources(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, BLUE)
        section_chip(slide, Inches(0.7), Inches(0.55), "LINKS", BLUE)
        add_textbox(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.55), "Resources", size=32, color=FG, bold=True)
        links = [
            ("Level 2 rice", "https://github.com/binnewbs/arch-hyprland"),
            ("Hyprland wiki", "https://wiki.hyprland.org"),
            ("Gruvbox", "https://github.com/morhetz/gruvbox"),
            ("Matugen", "https://github.com/InioX/matugen"),
            ("Waybar", "https://github.com/Alexays/Waybar"),
            ("Arch Wiki · Hyprland", "https://wiki.archlinux.org/title/Hyprland"),
        ]
        y = Inches(1.8)
        for name, url in links:
            panel(slide, Inches(0.7), y, Inches(12), Inches(0.7), BG_PANEL)
            add_textbox(slide, Inches(0.95), y + Inches(0.15), Inches(3.5), Inches(0.4), name, size=16, color=BLUE, bold=True)
            add_textbox(slide, Inches(4.6), y + Inches(0.15), Inches(7.8), Inches(0.4), url, size=15, color=FG_DIM)
            y += Inches(0.8)
        footer(slide, n, total)

    def end_slide(slide, n, total):
        fill_slide(slide)
        accent_bar(slide, ORANGE, height=Inches(0.18))
        accent_bar(slide, AQUA, top=Inches(0.18), height=Inches(0.06))
        add_textbox(slide, Inches(0.9), Inches(2.3), Inches(11.5), Inches(0.5), "NOW GO RICE", size=18, color=ORANGE, bold=True)
        add_textbox(slide, Inches(0.9), Inches(2.9), Inches(11.5), Inches(1.0), "Pick a level. Ship a desktop.", size=40, color=FG, bold=True)
        add_textbox(
            slide,
            Inches(0.9),
            Inches(4.1),
            Inches(11),
            Inches(1.0),
            "Level 1 — design with AI until it feels like yours.\nLevel 2 — start from arch-hyprland, then fork the look.",
            size=18,
            color=FG_DIM,
        )
        add_textbox(
            slide,
            Inches(0.9),
            Inches(5.5),
            Inches(11),
            Inches(0.5),
            "Import this .pptx into Google Slides: File → Import slides → Upload",
            size=15,
            color=GRAY,
        )
        footer(slide, n, total)

    builders = [
        title_slide,
        what_is_rice,
        two_levels,
        prereqs,
        stack,
        palette,
        level1_divider,
        level1_steps,
        level1_packages,
        level1_ai_prompts,
        level2_divider,
        level2_look,
        level2_install,
        level2_matugen,
        level2_layout,
        compare,
        polish,
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
