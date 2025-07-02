# # Version corrigée du script catalogue.py
# import os
# from PIL import Image
# import docx
# from docx.shared import Inches
# import pandas as pd
# from docx.oxml import parse_xml
# from docx.oxml.ns import nsdecls
# from docx.enum.text import WD_ALIGN_PARAGRAPH

# # === Paramètres ===
# root_path = r"F:\projet_img"

# # === Extraction des images avec structure à 3 niveaux ===
# def process_strate(strate_path):
#     images = []
#     print(f"  📁 Traitement du dossier: {strate_path}")
    
#     # Vérifier si le dossier existe
#     if not os.path.exists(strate_path):
#         print(f"  ❌ Dossier inexistant: {strate_path}")
#         return images
    
#     # Parcourir les dossiers de produits
#     for product_dir in os.listdir(strate_path):
#         product_path = os.path.join(strate_path, product_dir)
#         if not os.path.isdir(product_path):
#             continue
            
#         print(f"    📂 Produit: {product_dir}")
        
#         # Parcourir les dossiers d'unités
#         for unit_dir in os.listdir(product_path):
#             unit_path = os.path.join(product_path, unit_dir)
#             if not os.path.isdir(unit_path):
#                 continue
                
#             print(f"      📦 Unité: {unit_dir}")
            
#             # Rechercher directement les images dans le dossier unité
#             best_img = None
#             best_size = 0
#             images_found = 0
            
#             for file in os.listdir(unit_path):
#                 if file.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff')):
#                     img_path = os.path.join(unit_path, file)
#                     try:
#                         file_size = os.path.getsize(img_path)
#                         images_found += 1
#                         if file_size > best_size:
#                             best_size = file_size
#                             best_img = img_path
#                         print(f"        🖼️ Image trouvée: {file} ({file_size} bytes)")
#                     except Exception as e:
#                         print(f"        ❌ Erreur avec {file}: {e}")
#                         continue

#             if best_img:
#                 images.append({
#                     'path': best_img,
#                     'libelle_groupe': product_dir.strip(),
#                     'libelle_pdt': product_dir.strip(),  # Nom du produit
#                     'libelle_unite': unit_dir.strip()    # Unité
#                 })
#                 print(f"        ✅ Meilleure image sélectionnée: {os.path.basename(best_img)}")
#             else:
#                 print(f"        ⚠️ Aucune image trouvée dans {unit_dir}")
    
#     print(f"  📊 Total images trouvées pour cette strate: {len(images)}")
#     return images

# # === Fonction pour grouper les images par produit ===
# def group_images_by_product(images_list):
#     """Groupe les images par produit pour créer des pages séparées"""
#     grouped = {}
#     for img in images_list:
#         product = img['libelle_groupe']
#         if product not in grouped:
#             grouped[product] = []
#         grouped[product].append(img)
#     return grouped

# # === Redimensionnement d'image ===
# def resize_image_to_fixed_size(img_path, target_width_cm=7, target_height_cm=5):
#     """Redimensionne l'image en gardant les proportions"""
#     try:
#         with Image.open(img_path) as img:
#             # Convertir cm en pixels (approximation 96 DPI)
#             target_width_px = int(target_width_cm * 37.8)  # 96 DPI
#             target_height_px = int(target_height_cm * 37.8)
            
#             # Calculer les nouvelles dimensions en gardant les proportions
#             img_ratio = img.width / img.height
#             target_ratio = target_width_px / target_height_px
            
#             if img_ratio > target_ratio:
#                 # Image plus large - ajuster par la largeur
#                 new_width = target_width_px
#                 new_height = int(target_width_px / img_ratio)
#             else:
#                 # Image plus haute - ajuster par la hauteur
#                 new_height = target_height_px
#                 new_width = int(target_height_px * img_ratio)
            
#             # Redimensionner avec une bonne qualité
#             img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
#             # Sauvegarder temporairement (ou créer une copie)
#             temp_path = img_path.replace('.jpg', '_resized.jpg').replace('.png', '_resized.png')
#             img_resized.save(temp_path, quality=95)
#             return temp_path
#     except Exception as e:
#         print(f"    ❌ Erreur redimensionnement {img_path}: {e}")
#         return img_path  # Retourner l'original en cas d'erreur

# # === Créer un tableau sans bordures visibles ===
# def set_no_borders(table):
#     """Supprime les bordures du tableau"""
#     tbl = table._tbl
#     for row in tbl.tr_lst:
#         for cell in row.tc_lst:
#             tcPr = cell.tcPr
#             tcBorders = parse_xml(r'<w:tcBorders %s><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>' % nsdecls('w'))
#             tcPr.append(tcBorders)

# # === Création du catalogue Word ===
# def create_word_catalog(images_list, strate_name, output_filename):
#     """Crée le catalogue Word avec les images groupées par produit"""
#     if not images_list:
#         print(f"  ⚠️ Aucune image à traiter pour {strate_name}")
#         return
    
#     doc = docx.Document()

#     # Marges étroites
#     sections = doc.sections
#     for section in sections:
#         section.left_margin = Inches(0.5)
#         section.right_margin = Inches(0.5)
#         section.top_margin = Inches(0.5)
#         section.bottom_margin = Inches(0.5)

#     # Grouper les images par produit
#     grouped_images = group_images_by_product(images_list)
    
#     print(f"  📑 Création du catalogue avec {len(grouped_images)} groupes de produits")
    
#     # 6 images par page (3 lignes x 2 colonnes)
#     rows_per_page = 3
#     cols_per_page = 2
#     images_per_page = rows_per_page * cols_per_page

#     page_count = 0
    
#     # Pour chaque groupe de produits
#     for product_name, product_images in grouped_images.items():
#         print(f"    📄 Traitement du produit: {product_name} ({len(product_images)} images)")
        
#         # Traiter les images par pages de 6
#         for i in range(0, len(product_images), images_per_page):
#             page_images = product_images[i:i + images_per_page]
#             page_count += 1
            
#             # Titre de la page
#             title_paragraph = doc.add_paragraph()
#             title_run = title_paragraph.add_run(f'{strate_name} - {product_name}')
#             title_run.font.size = docx.shared.Pt(16)
#             title_run.bold = True
#             title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
#             # Créer le tableau
#             table = doc.add_table(rows=rows_per_page, cols=cols_per_page)
#             table.style = 'Table Grid'
#             set_no_borders(table)
            
#             # Remplir le tableau
#             for row in range(rows_per_page):
#                 for col in range(cols_per_page):
#                     idx = row * cols_per_page + col
#                     if idx < len(page_images):
#                         img_info = page_images[idx]
#                         cell = table.cell(row, col)
                        
#                         # Nettoyer la cellule
#                         cell.text = ''
                        
#                         # Ajouter le titre
#                         title_para = cell.add_paragraph()
#                         title_para.text = f'{img_info["libelle_pdt"]} - {img_info["libelle_unite"]}'
#                         title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
#                         # Redimensionner et ajouter l'image
#                         try:
#                             resized_img_path = resize_image_to_fixed_size(img_info['path'])
#                             img_para = cell.add_paragraph()
#                             img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
#                             img_run = img_para.add_run()
#                             img_run.add_picture(resized_img_path, width=Inches(2.5), height=Inches(1.8))
#                             print(f"      ✅ Image ajoutée: {os.path.basename(img_info['path'])}")
#                         except Exception as e:
#                             print(f"      ❌ Erreur ajout image {img_info['path']}: {e}")
#                             # Ajouter un texte de remplacement
#                             error_para = cell.add_paragraph()
#                             error_para.text = "[Image non disponible]"
#                             error_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
#             # Saut de page sauf pour la dernière page
#             if i + images_per_page < len(product_images) or product_name != list(grouped_images.keys())[-1]:
#                 doc.add_page_break()
    
#     # Sauvegarder le document
#     try:
#         doc.save(output_filename)
#         print(f"  ✅ Catalogue sauvegardé: {output_filename} ({page_count} pages)")
#     except Exception as e:
#         print(f"  ❌ Erreur sauvegarde: {e}")

# # === Liste des strates ===
# strates = [
#     "101_KOLDA"
#     # "102_VELINGARA",
#     # "103_MEDINA YORO FOULAH",
#     # "111_MATAM",
#     # "112_KANEL",
#     # "113_RANEROU",
#     # "11_DAKAR",
#     # "121_KAFFRINE",
#     # "122_BIRKELANE",
#     # "123_KOUNGHEUL",
#     # "124_MALEM HODDAR",
#     # "12_PIKINE",
#     # "131_KEDOUGOU",
#     # "132_SALEMATA",
#     # "133_SARAYA",
#     # "13_RUFISQUE",
#     # "141_SEDHIOU",
#     # "142_BOUNKILING",
#     # "143_GOUDOMP",
#     # "14_GUEDIAWAYE",
#     # "15_KEUR MASSAR",
#     # "21_BIGNONA",
#     # "22_OUSSOUYE",
#     # "23_ZIGUINCHOR",
#     # "31_BAMBEY",
#     # "32_DIOURBEL",
#     # "33_MBACKE",
#     # "41_DAGANA",
#     # "42_PODOR",
#     # "43_SAINT-LOUIS",
#     # "51_BAKEL",
#     # "52_TAMBACOUNDA",
#     # "53_GOUDIRY",
#     # "54_KOUMPENTOUM",
#     # "61_KAOLACK",
#     # "62_NIORO",
#     # "63_GUINGUINEO",
#     # "71_MBOUR",
#     # "72_THIES",
#     # "73_TIVAOUANE",
#     # "81_KEBEMER",
#     # "82_LINGUERE",
#     # "83_LOUGA",
#     # "91_FATICK",
#     # "92_FOUNDIOUGNE",
#     # "93_GOSSAS",
# ]

# # === Lancement ===
# print("🚀 Démarrage de la génération des catalogues...")
# print(f"📂 Répertoire racine: {root_path}")

# total_processed = 0
# total_errors = 0

# for strate in strates:
#     strate_path = os.path.join(root_path, strate)
#     if os.path.exists(strate_path):
#         print(f"\n▶️ Traitement de : {strate}")
#         try:
#             images = process_strate(strate_path)
            
#             if images:
#                 # Déterminer le chemin de sortie
#                 output_file = os.path.join(strate_path, f"{strate.replace(' ', '_')}_catalogue.docx")
                
#                 # Créer et enregistrer le fichier Word
#                 create_word_catalog(images, strate, output_file)
#                 total_processed += 1
#             else:
#                 print(f"  ⚠️ Aucune image trouvée dans {strate}")
                
#         except Exception as e:
#             print(f"  ❌ Erreur lors du traitement de {strate}: {e}")
#             total_errors += 1
#     else:
#         print(f"❌ Dossier introuvable : {strate_path}")
#         total_errors += 1

# print(f"\n✅ Génération terminée!")
# print(f"📊 Statistiques:")
# print(f"   - Strates traitées avec succès: {total_processed}")
# print(f"   - Erreurs rencontrées: {total_errors}")
# print(f"   - Total strates: {len(strates)}")
import os
import glob
from PIL import Image, UnidentifiedImageError
import docx
from docx.shared import Inches, Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from math import ceil
import traceback

# === Paramètres ===
root_path = r"F:\projet\projet_img" 
DPI = 96
LOG_FILE = "log_erreurs.txt"
# IMPORTANT: Assurez-vous que ce fichier existe dans le dossier du script
PLACEHOLDER_IMAGE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "placeholder.jpg")

# === Fonctions utilitaires (inchangées) ===
def process_strate(strate_path):
    # ... (code identique à la version précédente) ...
    images = []
    if not os.path.exists(strate_path):
        print(f"  ❌ Dossier inexistant: {strate_path}")
        return images
    for product_dir in os.listdir(strate_path):
        product_path = os.path.join(strate_path, product_dir)
        if not os.path.isdir(product_path): continue
        product_name = product_dir.strip()
        for unit_dir in os.listdir(product_path):
            unit_path = os.path.join(product_path, unit_dir)
            if not os.path.isdir(unit_path): continue
            best_img, best_size = None, 0
            for file in os.listdir(unit_path):
                if file.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff')):
                    img_path = os.path.join(unit_path, file)
                    try:
                        file_size = os.path.getsize(img_path)
                        if file_size > best_size:
                            best_size, best_img = file_size, img_path
                    except OSError: continue
            if best_img:
                images.append({
                    'path': best_img,
                    'libelle_produit': product_name,
                    'libelle_unite': unit_dir.strip()
                })
    return images

def group_images_by_product(images_list):
    # ... (code identique à la version précédente) ...
    grouped = {}
    for img in images_list:
        product = img['libelle_produit']
        if product not in grouped: grouped[product] = []
        grouped[product].append(img)
    return grouped

def log_error(path, message):
    # ... (code identique à la version précédente) ...
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{path} => {message}\n")

def cleanup_temp_images(directory):
    # ... (code identique à la version précédente) ...
    for file in glob.glob(os.path.join(directory, '**', '*_resized.*'), recursive=True):
        try: os.remove(file)
        except Exception as e: print(f"⚠️  Erreur suppression {file}: {e}")

def set_table_borders_invisible(table):
    # ... (code identique à la version précédente) ...
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = tblPr.first_child_found_in("w:tblBorders")
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border_el = OxmlElement(f'w:{border_name}')
        border_el.set(qn('w:val'), 'nil')
        tblBorders.append(border_el)


# === AMÉLIORATION MAJEURE: Redimensionnement avec Placeholder ===
def process_and_resize_image(img_path, target_width_inches):
    """
    Tente de redimensionner l'image. En cas d'échec, renvoie le chemin
    de l'image de remplacement (placeholder).
    """
    try:
        # Vérifie si le fichier image existe et n'est pas vide
        if not os.path.exists(img_path) or os.path.getsize(img_path) == 0:
            raise FileNotFoundError("Fichier inexistant ou vide.")

        with Image.open(img_path) as img:
            target_width_px = int(target_width_inches * DPI)
            img.thumbnail((target_width_px, 9999), Image.Resampling.LANCZOS)
            
            base, ext = os.path.splitext(img_path)
            save_ext = ".jpeg" if ext.lower() in [".jpg", ".jpeg"] else ".png"
            temp_path = f"{base}_resized{save_ext}"
            
            if img.mode != 'RGB':
                img = img.convert('RGB')
                
            img.save(temp_path, quality=90)
            return temp_path

    except (UnidentifiedImageError, FileNotFoundError, OSError) as e:
        print(f"    ❌ ERREUR IMAGE : {os.path.basename(img_path)}. Utilisation du placeholder. Raison: {e}")
        log_error(img_path, f"Image corrompue ou illisible ({e}). Remplacée par placeholder.")
        return PLACEHOLDER_IMAGE
    except Exception as e:
        print(f"    ❌ ERREUR INCONNUE sur {os.path.basename(img_path)}. Utilisation du placeholder. Raison: {e}")
        log_error(img_path, f"Erreur de traitement inattendue ({e}). Remplacée par placeholder.")
        return PLACEHOLDER_IMAGE

# === VERSION FINALE: Génération du catalogue ===
def create_word_catalog(images_list, strate_name, output_filename):
    if not images_list:
        print(f"  ⚠️ Aucune image à traiter pour {strate_name}")
        return

    # Vérifier si l'image de remplacement existe avant de commencer
    if not os.path.exists(PLACEHOLDER_IMAGE):
        print(f"FATAL: L'image de remplacement '{PLACEHOLDER_IMAGE}' est introuvable. Veuillez la créer.")
        return

    doc = docx.Document()
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Cm(1.5), Cm(1.5)
    section.left_margin, section.right_margin = Cm(1.5), Cm(1.5)

    grouped_images = group_images_by_product(images_list)
    print(f"  🎨 Création du catalogue design avec {len(grouped_images)} groupes de produits.")

    IMAGES_PER_ROW = 2
    GUTTER_WIDTH_CM = 1.2
    page_width_cm = section.page_width.cm - section.left_margin.cm - section.right_margin.cm
    cell_width_cm = (page_width_cm - GUTTER_WIDTH_CM * (IMAGES_PER_ROW - 1)) / IMAGES_PER_ROW
    image_width_cm = cell_width_cm - 0.2

    title = doc.add_heading(strate_name, level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size, title.runs[0].bold = Pt(20), True

    for product_name, product_images in grouped_images.items():
        sub_title = doc.add_paragraph()
        run = sub_title.add_run(product_name)
        run.bold, run.font.size = True, Pt(14)
        sub_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_title.paragraph_format.space_before, sub_title.paragraph_format.space_after = Pt(18), Pt(8)

        num_rows = ceil(len(product_images) / IMAGES_PER_ROW)
        table = doc.add_table(rows=num_rows, cols=IMAGES_PER_ROW)
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_table_borders_invisible(table)
        
        for col in table.columns:
            col.width = Cm(cell_width_cm)

        img_iterator = iter(product_images)
        for i in range(num_rows):
            for j in range(IMAGES_PER_ROW):
                try:
                    img_info = next(img_iterator)
                    cell = table.cell(i, j)
                    cell.text = ''
                    cell.vertical_alignment = docx.enum.table.WD_ALIGN_VERTICAL.TOP

                    caption_text = f"{img_info['libelle_produit']} - {img_info['libelle_unite']}"
                    cap_para = cell.add_paragraph()
                    cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = cap_para.add_run(caption_text)
                    run.bold, run.font.size = True, Pt(10)
                    cap_para.paragraph_format.space_after = Pt(6)

                    # --- Logique simplifiée ---
                    # Cette fonction renvoie TOUJOURS un chemin valide (soit l'image, soit le placeholder)
                    image_to_insert = process_and_resize_image(img_info['path'], target_width_inches=Cm(image_width_cm).inches)
                    
                    para_img = cell.add_paragraph()
                    para_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    para_img.add_run().add_picture(image_to_insert, width=Cm(image_width_cm))
                    
                    if image_to_insert != PLACEHOLDER_IMAGE:
                         print(f"      ✅ Image ajoutée : {os.path.basename(img_info['path'])}")

                except StopIteration:
                    pass
                except Exception:
                    log_error(img_info.get('path', 'inconnue'), f"Erreur insertion Word: {traceback.format_exc()}")
        
        doc.add_paragraph("").paragraph_format.space_after = Pt(6)

    try:
        doc.save(output_filename)
        print(f"\n  ✅ Catalogue design généré : {output_filename}")
    except Exception as e:
        print(f"  ❌ Erreur de sauvegarde du document : {e}")
        log_error(output_filename, f"Sauvegarde échouée : {e}")

    cleanup_temp_images(root_path)

# === Lancement ===
if __name__ == "__main__":
    if os.path.exists(LOG_FILE): os.remove(LOG_FILE)
        
    strates = ["101_KOLDA"]
    print("🚀 Démarrage de la génération des catalogues design...")
    print(f"📂 Répertoire racine: {root_path}")

    for strate in strates:
        strate_path = os.path.join(root_path, strate)
        print(f"\n▶️ Traitement de la strate : {strate}")
        try:
            images = process_strate(strate_path)
            if images:
                output_file = os.path.join(root_path, f"{strate}_catalogue_design_final.docx")
                create_word_catalog(images, strate, output_file)
            else:
                print(f"  ⚠️ Aucune image trouvée ou traitée dans {strate}")
        except Exception as e:
            print(f"  ❌ Erreur majeure lors du traitement de {strate}: {e}")
            log_error(strate_path, f"Erreur globale : {traceback.format_exc()}")

    print("\n\n✅ Génération terminée !")
    if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > 0:
        print(f"ℹ️  Un journal des erreurs a été créé ici : {os.path.abspath(LOG_FILE)}")