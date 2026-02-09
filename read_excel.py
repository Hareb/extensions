#!/usr/bin/env python3
"""Script pour lire le fichier Excel des succursales"""

import zipfile
import xml.etree.ElementTree as ET
import sys

def read_excel_simple(filename):
    """Lit un fichier Excel .xlsx simple"""

    with zipfile.ZipFile(filename, 'r') as zip_ref:
        # Lire les shared strings
        shared_strings = []
        try:
            with zip_ref.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                for si in root.findall('.//ss:si', ns):
                    t = si.find('.//ss:t', ns)
                    if t is not None:
                        shared_strings.append(t.text if t.text else '')
                    else:
                        shared_strings.append('')
        except KeyError:
            pass

        # Lire la feuille principale
        with zip_ref.open('xl/worksheets/sheet1.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

            rows_data = []
            for row in root.findall('.//ss:row', ns):
                row_data = []
                for cell in row.findall('.//ss:c', ns):
                    cell_type = cell.get('t')
                    value = cell.find('.//ss:v', ns)

                    if value is not None:
                        if cell_type == 's':  # Shared string
                            idx = int(value.text)
                            if idx < len(shared_strings):
                                row_data.append(shared_strings[idx])
                            else:
                                row_data.append('')
                        else:
                            row_data.append(value.text)
                    else:
                        row_data.append('')

                if any(row_data):  # Ne pas ajouter les lignes vides
                    rows_data.append(row_data)

            return rows_data

if __name__ == '__main__':
    filename = 'Succursales addresses.xlsx'
    rows = read_excel_simple(filename)

    print("=== CONTENU COMPLET DU FICHIER ===")
    print(f"Nombre de lignes: {len(rows)}\n")

    for i, row in enumerate(rows, 1):
        # Remplir les colonnes manquantes
        while len(row) < 3:
            row.append('')
        print(f"Ligne {i:2d}: [{row[0][:30]:30s}] | [{row[1][:50]:50s}] | [{row[2]:10s}]")

    print("\n=== FORMAT CSV ===")
    for row in rows:
        while len(row) < 3:
            row.append('')
        print(f'"{row[0]}","{row[1]}","{row[2]}"')
