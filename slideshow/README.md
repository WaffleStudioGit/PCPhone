# How to Rice Arch — Slideshow

Dark **Gruvbox** deck covering two ways to rice Arch + Hyprland:

1. **Level 1** — Build your own rice with AI (Cursor / chat)
2. **Level 2** — Use [binnewbs/arch-hyprland](https://github.com/binnewbs/arch-hyprland)

## File

| File | Use |
|------|-----|
| `How-to-Rice-Arch.pptx` | Main deck — import into **Google Slides** |
| `rice-arch.html` | Browser preview (same content, dark Gruvbox) |
| `assets/` | Screenshots from the arch-hyprland rice |
| `build_pptx.py` | Regenerates the `.pptx` |

## Open in Google Slides (Google Docs–compatible)

1. Upload `How-to-Rice-Arch.pptx` to [Google Drive](https://drive.google.com)
2. Right-click → **Open with → Google Slides**  
   *or* in an existing presentation: **File → Import slides → Upload** → select the `.pptx`
3. Optional: **File → Save as Google Slides** so it lives natively in Docs/Drive

Widescreen 16:9 (13.333″ × 7.5″). Fonts are Calibri/Consolas so Google Slides maps them cleanly.

## Rebuild the PPTX

```bash
python3 slideshow/build_pptx.py
```

Needs `python-pptx` (`pip install python-pptx`).
