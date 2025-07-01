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
from PIL import Image
import docx
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH

# === Paramètres ===
root_path = r"F:\projet\projet_img"
DPI = 96
CM_TO_PX = DPI * 0.393701
PLACEHOLDER_IMG = "placeholder.jpg"  # (optionnel)

# === Extraction des images avec structure à 3 niveaux ===
def process_strate(strate_path):
    images = []
    print(f"  📁 Traitement du dossier: {strate_path}")

    if not os.path.exists(strate_path):
        print(f"  ❌ Dossier inexistant: {strate_path}")
        return images

    for product_dir in os.listdir(strate_path):
        product_path = os.path.join(strate_path, product_dir)
        if not os.path.isdir(product_path):
            continue

        print(f"    📂 Produit: {product_dir}")

        for unit_dir in os.listdir(product_path):
            unit_path = os.path.join(product_path, unit_dir)
            if not os.path.isdir(unit_path):
                continue

            print(f"      📦 Unité: {unit_dir}")
            best_img = None
            best_size = 0

            for file in os.listdir(unit_path):
                if file.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff')):
                    img_path = os.path.join(unit_path, file)
                    try:
                        file_size = os.path.getsize(img_path)
                        if file_size > best_size:
                            best_size = file_size
                            best_img = img_path
                        print(f"        🖼️ Image trouvée: {file} ({file_size} bytes)")
                    except Exception as e:
                        print(f"        ❌ Erreur avec {file}: {e}")
                        continue

            if best_img:
                images.append({
                    'path': best_img,
                    'libelle_groupe': product_dir.strip(),
                    'libelle_pdt': product_dir.strip(),
                    'libelle_unite': unit_dir.strip()
                })
                print(f"        ✅ Meilleure image sélectionnée: {os.path.basename(best_img)}")
            else:
                print(f"        ⚠️ Aucune image trouvée dans {unit_dir}")

    print(f"  📊 Total images trouvées pour cette strate: {len(images)}")
    return images

# === Regroupement ===
def group_images_by_product(images_list):
    grouped = {}
    for img in images_list:
        product = img['libelle_groupe']
        if product not in grouped:
            grouped[product] = []
        grouped[product].append(img)
    return grouped

# === Redimensionnement d’image ===
def resize_image_to_fixed_size(img_path, target_width_cm=7, target_height_cm=5):
    try:
        with Image.open(img_path) as img:
            img.verify()  # test rapide
        with Image.open(img_path) as img:  # recharger pour traitement
            target_width_px = int(target_width_cm * CM_TO_PX)
            target_height_px = int(target_height_cm * CM_TO_PX)
            img_ratio = img.width / img.height
            target_ratio = target_width_px / target_height_px

            if img_ratio > target_ratio:
                new_width = target_width_px
                new_height = int(target_width_px / img_ratio)
            else:
                new_height = target_height_px
                new_width = int(target_height_px * img_ratio)

            img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            base, ext = os.path.splitext(img_path)
            temp_path = f"{base}_resized{ext}"
            img_resized.save(temp_path, quality=95)
            return temp_path
    except Exception as e:
        print(f"    ❌ Erreur redimensionnement {img_path}: {e}")
        log_error(img_path, f"Redimensionnement impossible : {e}")
        return None

# === Nettoyage des fichiers temporaires ===
def cleanup_temp_images(directory):
    for file in glob.glob(os.path.join(directory, '*_resized.*')):
        try:
            os.remove(file)
        except Exception as e:
            print(f"⚠️  Erreur suppression {file}: {e}")

# === Suppression des bordures ===
def set_no_borders(table):
    tbl = table._tbl
    for row in tbl.tr_lst:
        for cell in row.tc_lst:
            tcPr = cell.tcPr
            tcBorders = parse_xml(
                r'<w:tcBorders %s><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>' % nsdecls('w')
            )
            tcPr.append(tcBorders)

# === Journal des erreurs ===
def log_error(img_path, message):
    with open("log_erreurs.txt", "a", encoding="utf-8") as f:
        f.write(f"{img_path} => {message}\n")

# === Création du document Word ===
# === Création du catalogue Word ===
def create_word_catalog(images_list, strate_name, output_filename):
    if not images_list:
        print(f"  ⚠️ Aucune image à traiter pour {strate_name}")
        return

    doc = docx.Document()

    # Marges
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    grouped_images = group_images_by_product(images_list)
    print(f"  📑 Création du catalogue avec {len(grouped_images)} groupes")

    rows_per_page = 3
    cols_per_page = 2
    images_per_page = rows_per_page * cols_per_page

    for product_name, product_images in grouped_images.items():
        print(f"    🧾 Produit: {product_name} ({len(product_images)} images)")

        for i in range(0, len(product_images), images_per_page):
            page_images = product_images[i:i + images_per_page]

            title_paragraph = doc.add_paragraph()
            run = title_paragraph.add_run(f"{strate_name} - {product_name}")
            run.font.size = Pt(16)
            run.bold = True
            title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            table = doc.add_table(rows=rows_per_page, cols=cols_per_page)
            table.style = "Table Grid"
            set_no_borders(table)

            for row in range(rows_per_page):
                for col in range(cols_per_page):
                    idx = row * cols_per_page + col
                    if idx < len(page_images):
                        img_info = page_images[idx]
                        cell = table.cell(row, col)
                        cell.text = ""

                        # Titre
                        para = cell.add_paragraph()
                        para.text = f"{img_info['libelle_pdt']} - {img_info['libelle_unite']}"
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                        # Image
                        try:
                            resized_path = resize_image_to_fixed_size(img_info['path'])
                            if resized_path:
                                img_para = cell.add_paragraph()
                                img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                run = img_para.add_run()
                                run.add_picture(resized_path, width=Inches(2.5), height=Inches(1.8))
                                print(f"      ✅ Image ajoutée: {os.path.basename(img_info['path'])}")
                            else:
                                raise Exception("Image non redimensionnée")
                        except Exception as e:
                            print(f"      ❌ Erreur ajout image {img_info['path']}: {e}")
                            log_error(img_info['path'], f"Insertion échouée : {e}")
                            cell.add_paragraph("[Image indisponible]").alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Saut de page si nécessaire
            doc.add_page_break()

    try:
        doc.save(output_filename)
        print(f"  ✅ Catalogue sauvegardé : {output_filename}")
    except Exception as e:
        print(f"  ❌ Erreur sauvegarde : {e}")
        log_error(output_filename, f"Sauvegarde échouée : {e}")

# === Lancement global ===
if __name__ == "__main__":
    print("🚀 Lancement de la génération des catalogues")
    print(f"📁 Dossier racine : {root_path}")

    strates = [
        "101_KOLDA"
        
    ]

    total_ok = 0
    total_errors = 0

    for strate in strates:
        strate_path = os.path.join(root_path, strate)
        if os.path.exists(strate_path):
            print(f"\n▶️ Traitement de la strate : {strate}")
            try:
                images = process_strate(strate_path)
                if images:
                    output_file = os.path.join(strate_path, f"{strate}_catalogue.docx")
                    create_word_catalog(images, strate, output_file)
                    cleanup_temp_images(strate_path)
                    total_ok += 1
                else:
                    print(f"  ⚠️ Aucune image valide pour {strate}")
            except Exception as e:
                print(f"  ❌ Erreur lors du traitement de {strate} : {e}")
                log_error(strate, f"Erreur globale : {e}")
                total_errors += 1
        else:
            print(f"  ❌ Dossier inexistant : {strate_path}")
            total_errors += 1

    print("\n✅ Fin de la génération.")
    print(f"📊 Statistiques : {total_ok} traitées, {total_errors} erreurs.")
