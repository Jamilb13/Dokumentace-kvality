# AUTEL - Utilita pro Tvorbu Dokumentace Kvality (v1.2.0)

**Autor:** KBK (AUTEL, a.s.)  
**Verze:** v1.2.0 (Release 2026)  
**Git Repozitář:** [https://github.com/KBK/Utilita-Dokumentace-Kvality](https://github.com/KBK/Utilita-Dokumentace-Kvality)  

![AUTEL Logo](app_logo.png)

---

## 📌 Popis projektu

Aplikace **AUTEL - Utilita pro Tvorbu Dokumentace Kvality** slouží k automatizovanému zpracování, hromadnému přečíslování, razítkování a kompletaci PDF dokumentů kvality (*Quality Documentation*) podle podadresářové struktury.

Utilita se skládá ze samostatného rozhraní v Pythonu (CustomTkinter) a obousměrně propojeného sešitu MS Excel (`Dokumentace_Kvality.xlsx`).

---

## 🌟 Klíčové vlastnosti

- **Automatické rozpoznání adresářů & Unifikované označení**:
  - Dynamické vytváření číselných kódů podle hloubky vnořených podadresářů (`1.1.01`, `1.2.01` až `1.2.07`, `2.0.01`, `4.1.01`, `4.2.01`).
  - Podpora libovolné hloubky vnoření adresářů (5 i více úrovní) bez zbytečného duplikování prefixů.
- **Obousměrné živé propojení s MS Excel**:
  - Záložka **AUTEL**: Kompletní konfigurace (Cesty, Otáčení, Zachování struktury, Razítko, Barvy, Fonty).
  - Záložka **Seznam**: Seznam naskenovaných souborů s označením, počty stran a stavovými zápisy (zbarvení řádků podle stavu).
  - **Živý COM zápis**: Úpravy v GUI se okamžitě automaticky propisují přímo do otevřeného okna MS Excelu!
- **Razítkování & Formátování**:
  - Vypočítané označení se razítkuje do pravého horního rohu na každou stránku.
  - Podpora formátování fontu: **Tučné (Bold)**, *Kurzíva (Italic)* a <u>Podtržené</u>.
  - Vizuální vzorník 16 barev textu a pozadí (vč. průhledného pozadí).
  - Automatické otáčení stran naležato (Landscape) na výšku (Portrait).
- **Master PDF se Strukturovanými Záložkami**:
  - Sloučení všech processed PDF do jednoho kompletního dokumentu (`Kompletní_Dokumentace_Kvality.pdf`).
  - Vytvoření víceúrovňových PDF záložek (TOC), které **odkazují přímo na souřadnice razítka v pravém horním rohu stránky**.
- **Kompletní uživatelský návod přímo v aplikaci** a ve formátu PDF (`Uzivatelska_Prirucka_Utilita_Dokumentace_Kvality.pdf`).

---

## 🏗️ Architektura projektu

| Soubor | Popis |
| :--- | :--- |
| [`main.py`](file:///c:/Users/behalek/OneDrive%20-%20AUTEL,%20a.s/Antigravity/DK/main.py) | Hlavní spouštěcí bod aplikace. |
| [`gui.py`](file:///c:/Users/behalek/OneDrive%20-%20AUTEL,%20a.s/Antigravity/DK/gui.py) | Grafické uživatelské rozhraní (CustomTkinter), dialog návodu a živé event handlery. |
| [`excel_manager.py`](file:///c:/Users/behalek/OneDrive%20-%20AUTEL,%20a.s/Antigravity/DK/excel_manager.py) | Správa MS Excelu (`openpyxl` + `win32com.client`), čtení a zápis záložek AUTEL a Seznam. |
| [`scanner.py`](file:///c:/Users/behalek/OneDrive%20-%20AUTEL,%20a.s/Antigravity/DK/scanner.py) | Rekurzivní skenování adresářů, algoritmus výpočtu kódování `get_path_prefix()`. |
| [`pdf_processor.py`](file:///c:/Users/behalek/OneDrive%20-%20AUTEL,%20a.s/Antigravity/DK/pdf_processor.py) | Zpracování PDF pomocí PyMuPDF (`fitz`), otáčení stránek, kreslení razítek a generování TOC. |
| [`converter.py`](file:///c:/Users/behalek/OneDrive%20-%20AUTEL,%20a.s/Antigravity/DK/converter.py) | Konverze ne-PDF dokumentů na PDF. |
| [`Utilita_Dokumentace_Kvality.spec`](file:///c:/Users/behalek/OneDrive%20-%20AUTEL,%20a.s/Antigravity/DK/Utilita_Dokumentace_Kvality.spec) | Konfigurace PyInstalleru pro kompilaci do `.exe`. |
| [`Uzivatelska_Prirucka_Utilita_Dokumentace_Kvality.pdf`](file:///c:/Users/behalek/OneDrive%20-%20AUTEL,%20a.s/Antigravity/DK/Uzivatelska_Prirucka_Utilita_Dokumentace_Kvality.pdf) | Oficiální uživatelská příručka v PDF. |

---

## 🚀 Spuštění a Kompilace

### Spuštění zdrojového kódu (Python):
```bash
pip install customtkinter openpyxl pymupdf pywin32 pillow
python main.py
```

### Kompilace do spouštěcího souboru (.exe):
```bash
python -m PyInstaller Utilita_Dokumentace_Kvality.spec --noconfirm
```
Výsledný soubor `Utilita_Dokumentace_Kvality.exe` naleznete ve složce `dist/`.

---

## 📄 Licenční informace
**© 2026 AUTEL, a.s.**  
Autor: **KBK**  
Všechna práva vyhrazena.
