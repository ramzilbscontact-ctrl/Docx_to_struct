#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convertit le CSV en format Odoo SANS la colonne Tags
"""

import csv

INPUT = "/Users/ramzilbs/Desktop/radiance_crm/clients_fideles.csv"
OUTPUT = "/Users/ramzilbs/Desktop/radiance_crm/clients_odoo_final.csv"

print("🔄 Conversion pour Odoo (sans Tags)...")

try:
    # Lire le fichier original
    with open(INPUT, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        data = []
        
        for row in reader:
            nom = row.get('Nom', '').strip()
            prenom = row.get('Prénom', '').strip()
            phone = row.get('Téléphone', '').strip()
            nb_seances = row.get('Nombre de séances', '0')
            
            # Nom complet
            if prenom:
                name = f"{prenom} {nom}"
            else:
                name = nom
            
            if name:  # Ignorer si pas de nom
                data.append({
                    'Name': name,
                    'Phone': phone,
                    'Notes': f"Nombre de séances: {nb_seances}"
                })
    
    # Écrire le nouveau CSV
    with open(OUTPUT, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=['Name', 'Phone', 'Notes'])
        writer.writeheader()
        writer.writerows(data)
    
    print(f"✅ Fichier créé: {OUTPUT}")
    print(f"📊 {len(data)} clients")
    print("\n📋 Import Odoo:")
    print("1. Contacts → Favoris → Importer")
    print("2. Chargez: clients_odoo_final.csv")
    print("3. Mapping:")
    print("   Name → Nom")
    print("   Phone → Téléphone")
    print("   Notes → Notes")
    print("4. Importez !")
    
except FileNotFoundError:
    print(f"❌ Fichier introuvable: {INPUT}")
    print("Lancez d'abord: python3 main.py")