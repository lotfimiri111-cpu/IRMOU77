"""
Theme System — مذكرتي Pro v17
All color palettes defined in one place. Immutable after import.
"""
from __future__ import annotations
from dataclasses import dataclass
from pptx.dml.color import RGBColor


def _rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


@dataclass(frozen=True)
class Theme:
    name: str
    family: str

    # Main palette
    bg: str           # Background (dark)
    bg2: str          # Background secondary
    accent: str       # Primary accent (e.g. gold)
    accent2: str      # Secondary accent
    text_light: str   # Text on dark bg
    text_dark: str    # Text on light bg
    card: str         # Card background
    muted: str        # Muted/subtle color

    # Gradient stops
    grad1: str
    grad2: str
    accent_grad1: str
    accent_grad2: str

    def rgb(self, attr: str) -> RGBColor:
        return _rgb(getattr(self, attr))

    @property
    def bg_rgb(self): return _rgb(self.bg)
    @property
    def bg2_rgb(self): return _rgb(self.bg2)
    @property
    def accent_rgb(self): return _rgb(self.accent)
    @property
    def accent2_rgb(self): return _rgb(self.accent2)
    @property
    def text_light_rgb(self): return _rgb(self.text_light)
    @property
    def text_dark_rgb(self): return _rgb(self.text_dark)
    @property
    def card_rgb(self): return _rgb(self.card)
    @property
    def muted_rgb(self): return _rgb(self.muted)


THEMES: dict[str, Theme] = {
    "navy_gold": Theme(
        name="navy_gold", family="Navy & Gold",
        bg="#07172F", bg2="#0D2347",
        accent="#C6A03C", accent2="#E8C96A",
        text_light="#F0E8D8", text_dark="#07172F",
        card="#122040", muted="#8A9BB5",
        grad1="#07172F", grad2="#1A3A6E",
        accent_grad1="#C6A03C", accent_grad2="#E8C96A",
    ),
    "dark_teal": Theme(
        name="dark_teal", family="Dark Teal",
        bg="#0A1F1C", bg2="#0F2E29",
        accent="#2EBFA5", accent2="#5DD4BE",
        text_light="#E8F5F3", text_dark="#0A1F1C",
        card="#122820", muted="#7BAAA4",
        grad1="#0A1F1C", grad2="#1A4A42",
        accent_grad1="#2EBFA5", accent_grad2="#5DD4BE",
    ),
    "burgundy": Theme(
        name="burgundy", family="Burgundy",
        bg="#1A0A0F", bg2="#2D1018",
        accent="#C0392B", accent2="#E74C3C",
        text_light="#F5E8E8", text_dark="#1A0A0F",
        card="#2A1015", muted="#A07878",
        grad1="#1A0A0F", grad2="#3D1520",
        accent_grad1="#C0392B", accent_grad2="#E74C3C",
    ),
    "forest": Theme(
        name="forest", family="Forest",
        bg="#0D1F0D", bg2="#152815",
        accent="#27AE60", accent2="#58D68D",
        text_light="#E8F5E8", text_dark="#0D1F0D",
        card="#152215", muted="#7AAE7A",
        grad1="#0D1F0D", grad2="#1E4020",
        accent_grad1="#27AE60", accent_grad2="#58D68D",
    ),
    "midnight_purple": Theme(
        name="midnight_purple", family="Midnight Purple",
        bg="#0F0A1F", bg2="#1A1030",
        accent="#8E44AD", accent2="#BB8FCE",
        text_light="#F0E8FF", text_dark="#0F0A1F",
        card="#1A1028", muted="#8A7AAA",
        grad1="#0F0A1F", grad2="#2C1A50",
        accent_grad1="#8E44AD", accent_grad2="#BB8FCE",
    ),
    "charcoal_orange": Theme(
        name="charcoal_orange", family="Charcoal Orange",
        bg="#1A1A1A", bg2="#282828",
        accent="#E67E22", accent2="#F39C12",
        text_light="#F5F0E8", text_dark="#1A1A1A",
        card="#222222", muted="#909090",
        grad1="#1A1A1A", grad2="#3A3A3A",
        accent_grad1="#E67E22", accent_grad2="#F39C12",
    ),
    "ice_blue": Theme(
        name="ice_blue", family="Ice Blue",
        bg="#0A1520", bg2="#0F2035",
        accent="#3498DB", accent2="#85C1E9",
        text_light="#E8F4FD", text_dark="#0A1520",
        card="#122030", muted="#6A9EB5",
        grad1="#0A1520", grad2="#1A3A5A",
        accent_grad1="#3498DB", accent_grad2="#85C1E9",
    ),
    "sand_gold": Theme(
        name="sand_gold", family="Sand Gold",
        bg="#1F1505", bg2="#2E2008",
        accent="#D4AC0D", accent2="#F1C40F",
        text_light="#FDF5E6", text_dark="#1F1505",
        card="#28180A", muted="#A0906A",
        grad1="#1F1505", grad2="#3D2A08",
        accent_grad1="#D4AC0D", accent_grad2="#F1C40F",
    ),
    "slate_crimson": Theme(
        name="slate_crimson", family="Slate Crimson",
        bg="#1A1A2E", bg2="#16213E",
        accent="#E94560", accent2="#F07070",
        text_light="#F0F0F8", text_dark="#1A1A2E",
        card="#1E2040", muted="#8A8AB0",
        grad1="#1A1A2E", grad2="#0F3460",
        accent_grad1="#E94560", accent_grad2="#F07070",
    ),
    "noir": Theme(
        name="noir", family="Noir",
        bg="#0D0D0D", bg2="#1A1A1A",
        accent="#C9B99A", accent2="#E8D5B5",
        text_light="#F5F5F5", text_dark="#0D0D0D",
        card="#1A1A1A", muted="#888888",
        grad1="#0D0D0D", grad2="#2A2A2A",
        accent_grad1="#C9B99A", accent_grad2="#E8D5B5",
    ),
    "atlas": Theme(
        name="atlas", family="Atlas",
        bg="#0A0F1E", bg2="#101828",
        accent="#00BCD4", accent2="#4DD0E1",
        text_light="#E8F8FF", text_dark="#0A0F1E",
        card="#121F35", muted="#6A9BAB",
        grad1="#0A0F1E", grad2="#1A2F4E",
        accent_grad1="#00BCD4", accent_grad2="#4DD0E1",
    ),
    "sakura": Theme(
        name="sakura", family="Sakura",
        bg="#1A0A14", bg2="#2A1020",
        accent="#E91E8C", accent2="#F06EB5",
        text_light="#FFE8F5", text_dark="#1A0A14",
        card="#22102A", muted="#A07090",
        grad1="#1A0A14", grad2="#3A1530",
        accent_grad1="#E91E8C", accent_grad2="#F06EB5",
    ),


    # ── Canva-exclusive themes ──────────────────────────────────────────
    "canva_ocean": Theme(
        name="canva_ocean", family="Canva Ocean",
        bg="#030D1A", bg2="#051828",
        accent="#00C2FF", accent2="#38D9FF",
        text_light="#E0F7FF", text_dark="#030D1A",
        card="#061525", muted="#5A98B0",
        grad1="#030D1A", grad2="#0A2A42",
        accent_grad1="#00C2FF", accent_grad2="#38D9FF",
    ),
    "canva_sunset": Theme(
        name="canva_sunset", family="Canva Sunset",
        bg="#1A0A00", bg2="#2E1400",
        accent="#FF6B35", accent2="#FF9A5C",
        text_light="#FFF0E8", text_dark="#1A0A00",
        card="#251008", muted="#A06040",
        grad1="#1A0A00", grad2="#3D1800",
        accent_grad1="#FF6B35", accent_grad2="#FFC080",
    ),
    "canva_violet": Theme(
        name="canva_violet", family="Canva Violet",
        bg="#0D0818", bg2="#180F2E",
        accent="#A855F7", accent2="#C084FC",
        text_light="#F3E8FF", text_dark="#0D0818",
        card="#160C25", muted="#7A5A9A",
        grad1="#0D0818", grad2="#28154A",
        accent_grad1="#A855F7", accent_grad2="#C084FC",
    ),
    "canva_mint": Theme(
        name="canva_mint", family="Canva Mint",
        bg="#011A12", bg2="#042918",
        accent="#10B981", accent2="#34D399",
        text_light="#ECFDF5", text_dark="#011A12",
        card="#062018", muted="#4A8A6A",
        grad1="#011A12", grad2="#093D28",
        accent_grad1="#10B981", accent_grad2="#34D399",
    ),
    "canva_rose": Theme(
        name="canva_rose", family="Canva Rose",
        bg="#18050E", bg2="#2A0A18",
        accent="#F43F5E", accent2="#FB7185",
        text_light="#FFF1F2", text_dark="#18050E",
        card="#220812", muted="#8A4A5A",
        grad1="#18050E", grad2="#3A0F20",
        accent_grad1="#F43F5E", accent_grad2="#FB7185",
    ),
    "canva_amber": Theme(
        name="canva_amber", family="Canva Amber",
        bg="#160E00", bg2="#241500",
        accent="#F59E0B", accent2="#FCD34D",
        text_light="#FFFBEB", text_dark="#160E00",
        card="#201200", muted="#907030",
        grad1="#160E00", grad2="#3A2200",
        accent_grad1="#F59E0B", accent_grad2="#FCD34D",
    ),
    "canva_indigo": Theme(
        name="canva_indigo", family="Canva Indigo",
        bg="#04061A", bg2="#080E2E",
        accent="#6366F1", accent2="#818CF8",
        text_light="#EEF2FF", text_dark="#04061A",
        card="#0A0E25", muted="#5060A0",
        grad1="#04061A", grad2="#12184A",
        accent_grad1="#6366F1", accent_grad2="#818CF8",
    ),
    "canva_slate": Theme(
        name="canva_slate", family="Canva Slate",
        bg="#080C14", bg2="#101828",
        accent="#64748B", accent2="#94A3B8",
        text_light="#F1F5F9", text_dark="#080C14",
        card="#101520", muted="#4A5A72",
        grad1="#080C14", grad2="#1A2540",
        accent_grad1="#64748B", accent_grad2="#94A3B8",
    ),
    # ── Classic-exclusive (academic official) themes ───────────────────
    "classic_royal": Theme(
        name="classic_royal", family="Classic Royal Blue",
        bg="#001233", bg2="#012A72",
        accent="#B8860B", accent2="#DAA520",
        text_light="#F0F4FF", text_dark="#001233",
        card="#001840", muted="#607090",
        grad1="#001233", grad2="#023490",
        accent_grad1="#B8860B", accent_grad2="#DAA520",
    ),
    "classic_olive": Theme(
        name="classic_olive", family="Classic Olive Academic",
        bg="#0F1A00", bg2="#1E3300",
        accent="#8B9E2A", accent2="#B5C84E",
        text_light="#F4F8E8", text_dark="#0F1A00",
        card="#182500", muted="#5A7030",
        grad1="#0F1A00", grad2="#2E4E00",
        accent_grad1="#8B9E2A", accent_grad2="#B5C84E",
    ),
    "classic_maroon": Theme(
        name="classic_maroon", family="Classic Maroon",
        bg="#1A0000", bg2="#3A0010",
        accent="#C0963C", accent2="#E0B860",
        text_light="#FFF5F5", text_dark="#1A0000",
        card="#280005", muted="#905A5A",
        grad1="#1A0000", grad2="#4A0015",
        accent_grad1="#C0963C", accent_grad2="#E0B860",
    ),
    "classic_slate_blue": Theme(
        name="classic_slate_blue", family="Classic Slate Blue",
        bg="#0F1622", bg2="#1A2540",
        accent="#5B8DB8", accent2="#7EADCF",
        text_light="#EEF3F8", text_dark="#0F1622",
        card="#161E35", muted="#506080",
        grad1="#0F1622", grad2="#1E3255",
        accent_grad1="#5B8DB8", accent_grad2="#7EADCF",
    ),
    # ── Premium-exclusive themes ────────────────────────────────────────
    "premium_obsidian": Theme(
        name="premium_obsidian", family="Premium Obsidian",
        bg="#050508", bg2="#0C0C14",
        accent="#E8D5A3", accent2="#F5EAC8",
        text_light="#FAF8F5", text_dark="#050508",
        card="#0E0E18", muted="#6A6A80",
        grad1="#050508", grad2="#141420",
        accent_grad1="#E8D5A3", accent_grad2="#F5EAC8",
    ),
    "premium_royal_night": Theme(
        name="premium_royal_night", family="Premium Royal Night",
        bg="#000A1A", bg2="#001030",
        accent="#4FC3F7", accent2="#81D4FA",
        text_light="#E8F5FF", text_dark="#000A1A",
        card="#000E22", muted="#3A6080",
        grad1="#000A1A", grad2="#001840",
        accent_grad1="#4FC3F7", accent_grad2="#81D4FA",
    ),
    "premium_emerald": Theme(
        name="premium_emerald", family="Premium Emerald",
        bg="#020F08", bg2="#041A10",
        accent="#50FA7B", accent2="#80FFAA",
        text_light="#EDFFF4", text_dark="#020F08",
        card="#052010", muted="#2A5A3A",
        grad1="#020F08", grad2="#082818",
        accent_grad1="#50FA7B", accent_grad2="#80FFAA",
    ),
    "premium_copper": Theme(
        name="premium_copper", family="Premium Copper",
        bg="#0E0800", bg2="#1C1000",
        accent="#B87333", accent2="#D4924A",
        text_light="#FFF8F0", text_dark="#0E0800",
        card="#180C00", muted="#7A5030",
        grad1="#0E0800", grad2="#2A1800",
        accent_grad1="#B87333", accent_grad2="#D4924A",
    ),

}
def get_theme(name: str) -> Theme:
    return THEMES.get(name, THEMES["navy_gold"])
