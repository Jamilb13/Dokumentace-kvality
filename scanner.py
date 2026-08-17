import os
import fitz
from converter import ensure_pdf

SUPPORTED_EXTENSIONS = ['.pdf', '.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.webp', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt']

import re

def extract_folder_number(folder_name):
    """Vytáhne ze jména složky úvodní číselný kód (např. '1.1', '4.1.1', '2.0', '1')."""
    folder_name = str(folder_name).strip()
    m = re.match(r'^(\d+(?:\.\d+)*)', folder_name)
    if m:
        return m.group(1)
    m2 = re.search(r'(\d+(?:\.\d+)*)', folder_name)
    if m2:
        return m2.group(1)
    return ""

def get_path_prefix(parts):
    """Sestaví číselný kód spojením čísel všech podadresářů v cestě pro jakoukoliv hloubku vnoření."""
    if not parts or parts == ["."]:
        return "1.0"
        
    num_parts = []
    for idx, p in enumerate(parts, start=1):
        num = extract_folder_number(p)
        if not num:
            num = str(idx)
        
        if not num_parts:
            num_parts.append(num)
        else:
            prev_prefix = ".".join(num_parts)
            if num.startswith(prev_prefix + "."):
                num_parts = [num]
            elif num == prev_prefix:
                pass
            else:
                num_parts.append(num)
                
    return ".".join(num_parts)

def scan_directory(source_dir, check_blank_pages=True, target_dir=None, error_dir=None, existing_map=None):
    """
    Rekurzivně naskenuje složku source_dir a vrátí seznam slovníků pro záložku Seznam.
    Pokud existing_map obsahuje nastavené Označení pro daný soubor z Excelu, zachová ho!
    Automaticky ignoruje cílový adresář (target_dir) a chybový adresář (error_dir).
    """
    if not os.path.exists(source_dir):
        raise FileNotFoundError(f"Zdrojová složka neexistuje: {source_dir}")
        
    exclude_paths = set()
    if target_dir:
        exclude_paths.add(os.path.abspath(target_dir).lower())
    if error_dir:
        exclude_paths.add(os.path.abspath(error_dir).lower())

    found_files = []
    
    for root, dirs, files in os.walk(source_dir):
        # Vyloučit cílové a chybové složky z rekurzivního skenování
        dirs_to_keep = []
        for d in dirs:
            abs_d = os.path.abspath(os.path.join(root, d)).lower()
            is_excluded = False
            for ex in exclude_paths:
                if abs_d == ex or abs_d.startswith(ex + os.sep):
                    is_excluded = True
                    break
            if not is_excluded:
                dirs_to_keep.append(d)
                
        dirs[:] = dirs_to_keep
        dirs.sort()
        files.sort()
        
        for file_name in files:
            ext = os.path.splitext(file_name)[1].lower()
            if ext in SUPPORTED_EXTENSIONS and not file_name.startswith("~$") and not file_name.endswith("_converted.pdf"):
                full_path = os.path.join(root, file_name)
                found_files.append((root, file_name, full_path))
                
    file_entries = []
    dir_counter = {}
    
    for root, file_name, full_path in found_files:
        rel_path = os.path.relpath(root, source_dir)
        parts = rel_path.split(os.sep) if rel_path != "." else ["."]
        prefix = get_path_prefix(parts)
            
        if prefix not in dir_counter:
            dir_counter[prefix] = 1
        else:
            dir_counter[prefix] += 1
            
        seq_num = f"{dir_counter[prefix]:02d}"
        default_oznaceni = f"{prefix}.{seq_num}"
        
        # Zachovat existující označení z Excelu pouze pokud NENÍ v chybné staré podobě!
        full_path_key = os.path.abspath(full_path).lower()
        full_oznaceni = default_oznaceni
        if existing_map:
            old_val = existing_map.get(full_path_key) or existing_map.get(full_path.lower())
            if old_val:
                # Pokud stará hodnota neobsahuje dostatek úrovní teček pro dané vnoření, nahradit ji
                expected_dots = prefix.count(".") + 1
                actual_dots = str(old_val).count(".")
                if actual_dots < expected_dots:
                    full_oznaceni = default_oznaceni
                else:
                    full_oznaceni = str(old_val).strip()
        
        page_count = 0
        status = "OK"
        blank_warnings = []
        
        try:
            target_pdf = ensure_pdf(full_path)
            doc = fitz.open(target_pdf)
            page_count = len(doc)
            
            if check_blank_pages and page_count > 0:
                for page_num in range(page_count):
                    page = doc[page_num]
                    text = page.get_text().strip()
                    images = page.get_images()
                    drawings = page.get_drawings()
                    
                    if len(text) == 0 and len(images) == 0 and len(drawings) == 0:
                        blank_warnings.append(str(page_num + 1))
                        
            doc.close()
            
            if page_count == 0:
                status = "CHYBA (0 stran)"
            elif blank_warnings:
                status = f"VAROVÁNÍ (Prázdná st. {', '.join(blank_warnings)})"
                
        except Exception as e:
            status = f"CHYBA ({str(e)})"
            page_count = 0
            
        base_name_no_ext = os.path.splitext(file_name)[0]
        novy_nazev = f"{full_oznaceni} - {base_name_no_ext}.pdf" if full_oznaceni else f"{base_name_no_ext}.pdf"
        
        file_entries.append({
            'oznaceni': full_oznaceni,
            'cesta': full_path,
            'nazev': file_name,
            'pocet_stran': page_count,
            'zapis': status,
            'novy_nazev': novy_nazev
        })
        
    return file_entries
