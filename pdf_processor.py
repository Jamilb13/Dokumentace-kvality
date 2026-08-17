import os
import shutil
import fitz
from converter import ensure_pdf, handle_protected_pdf
from excel_manager import safe_save_workbook, read_open_excel_via_com, write_to_open_excel_via_com, get_config_sheet

def parse_hex_color(hex_str, default_rgb=(0, 0, 0)):
    if not hex_str:
        return default_rgb
    hex_str = str(hex_str).strip().lstrip('#')
    if hex_str.upper() in ["TRANSPARENT", "NONE", "0", "FFFFFF00"]:
        return None
    if len(hex_str) == 6:
        try:
            r = int(hex_str[0:2], 16) / 255.0
            g = int(hex_str[2:4], 16) / 255.0
            b = int(hex_str[4:6], 16) / 255.0
            return (r, g, b)
        except ValueError:
            pass
    return default_rgb

def process_pdfs_from_excel(wb, excel_path=None, log_callback=None):
    """
    Provede kompletaci a razítkování podle nastavení v Excelu.
    Pokud je MS Excel otevřen, vyčte i zapíše data přímo přes COM API bez dočasných souborů!
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    com_config, com_entries = (None, None)
    if excel_path and os.path.exists(excel_path):
        com_config, com_entries = read_open_excel_via_com(excel_path)

    if com_config and com_entries:
        log("Čtu konfiguraci i seznam souborů ŽIVĚ přímo z otevřeného MS Excelu přes COM API...")
        config = com_config
        raw_file_list = com_entries
        using_com_live = True
    else:
        ws_config = get_config_sheet(wb)
        config = {}
        for r in range(2, ws_config.max_row + 1):
            k = ws_config.cell(r, 1).value
            v = ws_config.cell(r, 2).value
            if k:
                config[str(k).strip()] = str(v).strip() if v is not None else ""
        
        ws_seznam = wb["Seznam"]
        raw_file_list = []
        for r in range(2, ws_seznam.max_row + 1):
            ozn = ws_seznam.cell(r, 1).value
            path = ws_seznam.cell(r, 2).value
            name = ws_seznam.cell(r, 3).value
            pages = ws_seznam.cell(r, 4).value
            status = ws_seznam.cell(r, 5).value
            new_name = ws_seznam.cell(r, 6).value
            if path or name or ozn:
                raw_file_list.append({
                    'row_idx': r,
                    'oznaceni': str(ozn).strip() if ozn is not None else "",
                    'cesta': str(path).strip() if path is not None else "",
                    'nazev': str(name).strip() if name is not None else "",
                    'pocet_stran': int(pages) if pages is not None and str(pages).isdigit() else 0,
                    'zapis': str(status).strip() if status is not None else "",
                    'novy_nazev': str(new_name).strip() if new_name is not None else ""
                })
        using_com_live = False

    src_dir = config.get("Zdroj", "")
    dst_dir = config.get("Cíl", "")
    err_dir = config.get("Chybné soubory", "")
    overwrite = config.get("Přepsat existující soubory", "ANO").upper() == "ANO"
    purge_dst = config.get("Smazat cílový adresář", "NE").upper() == "ANO"
    rotate_portrait = config.get("Otáčet stránky na výšku", "ANO").upper() == "ANO"
    keep_structure = config.get("Zachovat adresářovou strukturu", "NE").upper() == "ANO"
    merge_single = config.get("Kompletovat do jednoho PDF", "NE").upper() == "ANO"
    
    font_size = float(config.get("Velikost fontu razítka", 10))
    is_bold = config.get("Tučné razítko", "NE").upper() == "ANO"
    is_italic = config.get("Kurzíva razítka", "NE").upper() == "ANO"
    is_underline = config.get("Podtržené razítko", "NE").upper() == "ANO"

    if is_bold and is_italic:
        fontname = "hebi"
    elif is_bold:
        fontname = "hebo"
    elif is_italic:
        fontname = "heit"
    else:
        fontname = "helv"

    text_color = parse_hex_color(config.get("Barva fontu razítka", "#000000"), (0, 0, 0))
    fill_color = parse_hex_color(config.get("Barva pozadí razítka", "#FFFFFF"), (1, 1, 1))
    opacity_pct = float(config.get("Průhlednost pozadí (%)", 100))
    fill_opacity = opacity_pct / 100.0

    if not dst_dir:
        raise ValueError("Cílový adresář není zadán v Konfiguraci.")

    os.makedirs(dst_dir, exist_ok=True)
    if err_dir:
        os.makedirs(err_dir, exist_ok=True)

    if purge_dst and os.path.exists(dst_dir):
        log(f"Čistím cílový adresář: {dst_dir}")
        abs_err = os.path.abspath(err_dir).lower() if err_dir else ""
        for item in os.listdir(dst_dir):
            item_p = os.path.join(dst_dir, item)
            abs_item = os.path.abspath(item_p).lower()
            if abs_err and (abs_item == abs_err or abs_item.startswith(abs_err + os.sep)):
                continue  # Přeskočit mazání chybové složky, pokud je uvnitř cíle!
            if os.path.isfile(item_p):
                try:
                    os.remove(item_p)
                except Exception as e:
                    log(f"Varování při mazání {item_p}: {e}")
            elif os.path.isdir(item_p):
                try:
                    shutil.rmtree(item_p)
                except Exception as e:
                    log(f"Varování při mazání složky {item_p}: {e}")

    processed_count = 0
    error_count = 0

    master_pdf = fitz.open() if merge_single else None
    toc_entries = []
    added_folders = set()

    for entry in raw_file_list:
        row_idx = entry.get('row_idx', 0)
        oznaceni = entry.get('oznaceni', '')
        full_path = entry.get('cesta', '')
        file_name = entry.get('nazev', '')
        novy_nazev = entry.get('novy_nazev', '')

        if not full_path or not os.path.exists(full_path):
            entry['zapis'] = "CHYBA (Nenalezen)"
            if not using_com_live and wb:
                wb["Seznam"].cell(row_idx, 5, value="CHYBA (Nenalezen)")
            error_count += 1
            continue

        if not novy_nazev:
            novy_nazev = f"{oznaceni} - {file_name}.pdf" if oznaceni else f"{file_name}.pdf"
            entry['novy_nazev'] = novy_nazev

        # Výpočet cílové cesty (buď přímo v Cíli, nebo v podadresářové struktuře podle zdroje)
        rel_dir = os.path.relpath(os.path.dirname(full_path), src_dir) if src_dir and os.path.exists(src_dir) else "."
        if keep_structure and rel_dir != ".":
            out_sub_dir = os.path.join(dst_dir, rel_dir)
            os.makedirs(out_sub_dir, exist_ok=True)
            target_out_path = os.path.join(out_sub_dir, novy_nazev)
        else:
            target_out_path = os.path.join(dst_dir, novy_nazev)

        if os.path.exists(target_out_path) and not overwrite:
            log(f"Přeskakuji (existuje): {novy_nazev}")
            entry['zapis'] = "OK (Přeskočeno)"
            if not using_com_live and wb:
                wb["Seznam"].cell(row_idx, 5, value="OK (Přeskočeno)")
            continue

        log(f"Zpracovávám [{oznaceni}]: {file_name}")

        try:
            # 1. Konverze ne-PDF na PDF v případě potřeby
            pdf_path = ensure_pdf(full_path)

            # 2. Odemčení / rasterizace chráněného PDF
            doc = handle_protected_pdf(pdf_path)

            if len(doc) == 0:
                raise ValueError("Dokument má 0 stran")

            # 3. Otáčení stránek a Razítkování na KAŽDOU stránku do Pravého Horního Rohu
            for page_idx in range(len(doc)):
                page = doc[page_idx]

                # Otáčení na výšku (Portrait) pokud je stránka naležato (Landscape)
                if rotate_portrait and page.rect.width > page.rect.height:
                    page.set_rotation((page.rotation + 90) % 360)

                # Výpočet pozice v pravém horním rohu
                margin_x = 15.0
                margin_y = 15.0
                stamp_text = str(oznaceni) if oznaceni else ""

                if stamp_text:
                    font = fitz.Font(fontname)
                    text_width = font.text_length(stamp_text, fontsize=font_size)
                    line_height = font_size * 1.2
                    padding = 4.0

                    rect_width = text_width + (padding * 2)
                    rect_height = line_height + (padding * 2)

                    rx2 = page.rect.width - margin_x
                    rx1 = rx2 - rect_width
                    ry1 = margin_y
                    ry2 = ry1 + rect_height

                    visual_rect = fitz.Rect(rx1, ry1, rx2, ry2)
                    text_pt_visual = fitz.Point(rx1 + padding, ry1 + padding + (font_size * 0.85))

                    if page.rotation != 0:
                        inv_M = ~page.rotation_matrix
                        draw_rect = visual_rect * inv_M
                        draw_text_pt = text_pt_visual * inv_M
                    else:
                        draw_rect = visual_rect
                        draw_text_pt = text_pt_visual

                    shape = page.new_shape()
                    shape.draw_rect(draw_rect)
                    shape.finish(
                        fill=fill_color,
                        color=fill_color,
                        fill_opacity=fill_opacity
                    )

                    if is_underline:
                        u_y_visual = ry1 + padding + (font_size * 0.85) + 1.5
                        u_pt1_visual = fitz.Point(rx1 + padding, u_y_visual)
                        u_pt2_visual = fitz.Point(rx1 + padding + text_width, u_y_visual)

                        if page.rotation != 0:
                            u_pt1 = u_pt1_visual * inv_M
                            u_pt2 = u_pt2_visual * inv_M
                        else:
                            u_pt1 = u_pt1_visual
                            u_pt2 = u_pt2_visual

                        shape.draw_line(u_pt1, u_pt2)
                        shape.finish(color=text_color, width=max(1.0, font_size * 0.08))

                    shape.commit()

                    if page.rotation != 0:
                        page.insert_text(
                            draw_text_pt,
                            stamp_text,
                            fontsize=font_size,
                            color=text_color,
                            fontname=fontname,
                            morph=(draw_text_pt, fitz.Matrix(page.rotation))
                        )
                    else:
                        page.insert_text(
                            draw_text_pt,
                            stamp_text,
                            fontsize=font_size,
                            color=text_color,
                            fontname=fontname
                        )

            doc.save(target_out_path)

            if merge_single:
                start_page_1based = len(master_pdf) + 1
                master_pdf.insert_pdf(doc)
                
                # Výpočet přesných souřadnic razítka na první straně pro přímý skok záložky
                p1 = doc[0]
                margin_x = 15.0
                margin_y = 15.0
                visual_stamp_pt = fitz.Point(p1.rect.width - margin_x, margin_y)
                if p1.rotation != 0:
                    stamp_target_pt = visual_stamp_pt * (~p1.rotation_matrix)
                else:
                    stamp_target_pt = visual_stamp_pt

                # Budování strukturovaných záložek (TOC/outlines) s přímým skokem na razítko označení
                if rel_dir and rel_dir != ".":
                    parts = rel_dir.split(os.sep)
                    for lvl in range(1, len(parts) + 1):
                        sub_path = os.sep.join(parts[:lvl])
                        if sub_path not in added_folders:
                            added_folders.add(sub_path)
                            toc_entries.append([
                                lvl, 
                                parts[lvl - 1], 
                                start_page_1based, 
                                {"kind": fitz.LINK_GOTO, "page": start_page_1based - 1, "to": stamp_target_pt}
                            ])
                    doc_lvl = len(parts) + 1
                else:
                    doc_lvl = 1
                    
                if novy_nazev.startswith(oznaceni):
                    doc_bookmark_title = novy_nazev
                elif oznaceni:
                    doc_bookmark_title = f"[{oznaceni}] {novy_nazev}"
                else:
                    doc_bookmark_title = novy_nazev

                toc_entries.append([
                    doc_lvl, 
                    doc_bookmark_title, 
                    start_page_1based, 
                    {"kind": fitz.LINK_GOTO, "page": start_page_1based - 1, "to": stamp_target_pt}
                ])

            doc.close()

            entry['zapis'] = "OK"
            if not using_com_live and wb:
                wb["Seznam"].cell(row_idx, 5, value="OK")
            processed_count += 1

        except Exception as e:
            log(f"CHYBA při zpracování {file_name}: {e}")
            entry['zapis'] = f"CHYBA ({str(e)})"
            if not using_com_live and wb:
                wb["Seznam"].cell(row_idx, 5, value=f"CHYBA ({str(e)})")
            error_count += 1
            if err_dir:
                try:
                    shutil.copy(full_path, os.path.join(err_dir, file_name))
                except Exception:
                    pass

    if merge_single and len(master_pdf) > 0:
        master_out = os.path.join(dst_dir, "Kompletní_Dokumentace_Kvality.pdf")
        log(f"Ukládám sloučené Master PDF se strukturovanými záložkami: {master_out}")
        master_pdf.set_toc(toc_entries)
        master_pdf.save(master_out)
        master_pdf.close()

    log(f"Zpracování dokončeno! Úspěšně: {processed_count}, Chyby: {error_count}")
    
    if excel_path:
        try:
            if using_com_live:
                log("Ukládám výsledné stavy ŽIVĚ přímo do MS Excelu přes COM API...")
                write_to_open_excel_via_com(excel_path, file_entries=raw_file_list, config_dict=config, only_update_status=True)
            else:
                saved, save_err = safe_save_workbook(wb, excel_path, file_entries=raw_file_list, config_dict=config, only_update_status=True)
                if not saved:
                    log(f"UPOZORNĚNÍ: {save_err}")
        except Exception as e:
            log(f"Varování při ukládání stavů v Excelu: {e}")

    return processed_count, error_count
