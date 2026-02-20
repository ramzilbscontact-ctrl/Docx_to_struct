#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script d'extraction et consolidation des clients fidèles depuis fichiers DOCX
Identifie les clients avec au moins 2 séances et fusionne les doublons
"""

import os
import re
import csv
from pathlib import Path
from typing import List, Dict, Tuple, Set
from datetime import datetime

# Imports des bibliothèques nécessaires
try:
    from docx2python import docx2python
    import pandas as pd
    from rapidfuzz import fuzz
    import dateparser
except ImportError as e:
    print(f"❌ Erreur: Une bibliothèque requise est manquante: {e}")
    print("\nInstallez les dépendances avec:")
    print("pip install --break-system-packages docx2python pandas rapidfuzz python-dateutil dateparser")
    exit(1)


# ============================================================================
# CONFIGURATION
# ============================================================================

INPUT_DIR = "/Users/ramzilbs/Desktop/radiance_crm/DOCX_SOURCE"
OUTPUT_FILE = "/Users/ramzilbs/Desktop/radiance_crm/clients_fideles.csv"
MIN_SESSIONS = 2  # Minimum de séances pour être considéré comme fidèle
FUZZY_THRESHOLD = 85  # Score minimum pour considérer deux noms comme similaires


# ============================================================================
# FONCTIONS UTILITAIRES
# ============================================================================

def normalize_phone(phone_str: str) -> str:
    """
    Normalise un numéro de téléphone en gardant uniquement les chiffres.
    Conserve les numéros avec au moins 9 chiffres.
    
    Args:
        phone_str: Chaîne contenant potentiellement un numéro de téléphone
        
    Returns:
        Numéro normalisé ou chaîne vide
    """
    if not phone_str:
        return ""
    
    # Extraire tous les chiffres
    digits = re.sub(r'\D', '', str(phone_str))
    
    # Garder seulement si au moins 9 chiffres
    return digits if len(digits) >= 9 else ""


def extract_phone_from_text(text: str) -> Tuple[str, str]:
    """
    Extrait le numéro de téléphone d'un texte et retourne le texte nettoyé.
    
    Args:
        text: Texte contenant potentiellement un nom et un téléphone
        
    Returns:
        Tuple (texte_sans_telephone, telephone_normalise)
    """
    # Patterns pour détecter les numéros de téléphone
    phone_patterns = [
        r'0\d{9}',  # 0XXXXXXXXX
        r'0\d{2}[\s\.-]?\d{2}[\s\.-]?\d{2}[\s\.-]?\d{2}[\s\.-]?\d{2}',  # 0X XX XX XX XX
        r'\+?\d{2,4}[\s\.-]?\d{2,4}[\s\.-]?\d{2,4}[\s\.-]?\d{2,4}',  # International
        r'\d{9,}',  # Au moins 9 chiffres consécutifs
    ]
    
    phone = ""
    clean_text = text
    
    for pattern in phone_patterns:
        match = re.search(pattern, text)
        if match:
            phone_candidate = normalize_phone(match.group(0))
            if phone_candidate:
                phone = phone_candidate
                # Supprimer le numéro du texte
                clean_text = re.sub(pattern, '', text).strip()
                break
    
    return clean_text, phone


def is_valid_name(text: str) -> bool:
    """
    Vérifie si un texte ressemble à un nom valide (pas une date ou un nombre).
    
    Args:
        text: Texte à vérifier
        
    Returns:
        True si c'est probablement un nom, False sinon
    """
    if not text or not isinstance(text, str):
        return False
    
    text = text.strip()
    
    # Rejeter si vide
    if not text:
        return False
    
    # Rejeter si c'est uniquement des chiffres et espaces (comme "22 02")
    if re.match(r'^[\d\s/\-\.]+$', text):
        return False
    
    # Rejeter si ça ressemble à une date (DD/MM, DD/MM/YY, etc.)
    if re.match(r'^\d{1,2}[/\-\.]\d{1,2}', text):
        return False
    
    # Rejeter si trop court (moins de 2 caractères)
    if len(text.strip()) < 2:
        return False
    
    # Accepter si ça contient au moins une lettre
    if re.search(r'[a-zA-ZÀ-ÿ]', text):
        return True
    
    return False


def parse_name(text: str) -> Tuple[str, str, str]:
    """
    Parse un texte pour extraire Nom, Prénom et Téléphone.
    Gère les formats: "Nom Prénom", "Prénom Nom", "Nom Prénom 0XXXXXXXXX"
    
    Args:
        text: Texte à parser
        
    Returns:
        Tuple (nom, prenom, telephone)
    """
    if not text or not isinstance(text, str):
        return "", "", ""
    
    # Nettoyer le texte
    text = text.strip()
    
    # Vérifier si c'est un nom valide
    if not is_valid_name(text):
        return "", "", ""
    
    # Extraire le téléphone d'abord
    text_without_phone, phone = extract_phone_from_text(text)
    
    # Nettoyer les caractères spéciaux et espaces multiples
    text_without_phone = re.sub(r'[^\w\s\-]', ' ', text_without_phone)
    text_without_phone = re.sub(r'\s+', ' ', text_without_phone).strip()
    
    if not text_without_phone:
        return "", "", phone
    
    # Vérifier à nouveau après nettoyage
    if not is_valid_name(text_without_phone):
        return "", "", ""
    
    # Séparer les mots
    parts = text_without_phone.split()
    
    if len(parts) == 0:
        return "", "", phone
    elif len(parts) == 1:
        # Un seul mot, considérer comme nom de famille
        return parts[0].title(), "", phone
    else:
        # Deux mots ou plus
        # Premier mot = Nom, Deuxième mot = Prénom
        nom = parts[0].title()
        prenom = parts[1].title()
        return nom, prenom, phone


def flatten_cell_content(cell_data) -> str:
    """
    Aplatit le contenu d'une cellule qui peut être une liste imbriquée.
    docx2python retourne des structures comme [[[['text']]]]
    
    Args:
        cell_data: Données de cellule (peut être liste, str, etc.)
        
    Returns:
        Texte aplati en une seule chaîne
    """
    if isinstance(cell_data, str):
        return cell_data.strip()
    elif isinstance(cell_data, list):
        # Récursivement aplatir
        result = []
        for item in cell_data:
            flattened = flatten_cell_content(item)
            if flattened:
                result.append(flattened)
        return ' '.join(result)
    else:
        return str(cell_data).strip() if cell_data else ""


def parse_dates(date_text: str) -> List[str]:
    """
    Parse une chaîne contenant une ou plusieurs dates séparées par virgules ou retours à la ligne.
    Filtre les dates aberrantes (avant 2000 ou après 2030).
    
    Args:
        date_text: Texte contenant des dates
        
    Returns:
        Liste de dates au format DD/MM/YYYY
    """
    if not date_text or not isinstance(date_text, str):
        return []
    
    # Séparer par virgules, points-virgules, et retours à la ligne
    separators = [',', ';', '\n', '\r']
    for sep in separators:
        date_text = date_text.replace(sep, '|')
    
    # Séparer et nettoyer
    date_parts = [d.strip() for d in date_text.split('|') if d.strip()]
    
    dates = []
    for date_str in date_parts:
        # Essayer de parser la date
        try:
            # Tentative avec dateparser
            parsed_date = dateparser.parse(
                date_str,
                settings={
                    'DATE_ORDER': 'DMY',  # Jour/Mois/Année (format français)
                    'PREFER_DAY_OF_MONTH': 'first',
                    'STRICT_PARSING': False
                }
            )
            if parsed_date:
                # Filtrer les dates aberrantes (avant 2000 ou après 2030)
                year = parsed_date.year
                if 2000 <= year <= 2030:
                    dates.append(parsed_date.strftime('%d/%m/%Y'))
            else:
                # Si dateparser échoue, essayer un pattern simple DD/MM/YYYY
                match = re.search(r'(\d{1,2})[/\.-](\d{1,2})[/\.-](\d{2,4})', date_str)
                if match:
                    day, month, year = match.groups()
                    # Ajouter le siècle si année sur 2 chiffres
                    if len(year) == 2:
                        year = '20' + year if int(year) < 50 else '19' + year
                    
                    # Filtrer les dates aberrantes
                    if 2000 <= int(year) <= 2030:
                        dates.append(f"{day.zfill(2)}/{month.zfill(2)}/{year}")
        except Exception:
            # Ignorer les dates qui ne peuvent pas être parsées
            continue
    
    return dates


def split_clients_in_cell(cell_text: str) -> List[str]:
    """
    Sépare plusieurs clients présents dans une même cellule.
    Détecte les séparateurs: retours à la ligne, points-virgules.
    
    Args:
        cell_text: Texte de la cellule
        
    Returns:
        Liste de textes (un par client)
    """
    if not cell_text:
        return []
    
    # Séparer par retours à la ligne et points-virgules
    clients = re.split(r'[\n\r;]+', cell_text)
    
    # Nettoyer et filtrer les entrées vides
    clients = [c.strip() for c in clients if c.strip()]
    
    return clients


# ============================================================================
# EXTRACTION DES DONNÉES
# ============================================================================

def extract_clients_from_docx(filepath: str) -> List[Dict]:
    """
    Extrait tous les clients et leurs séances depuis un fichier DOCX.
    
    Args:
        filepath: Chemin vers le fichier DOCX
        
    Returns:
        Liste de dictionnaires contenant les infos clients
    """
    print(f"\n📄 Traitement du fichier: {os.path.basename(filepath)}")
    
    clients = []
    
    try:
        # Ouvrir le document avec docx2python
        doc = docx2python(filepath)
        
        # Les tableaux sont dans doc.body
        for table in doc.body:
            if not table:
                continue
                
            # Parcourir chaque ligne du tableau
            for row_idx, row in enumerate(table):
                if not row or len(row) == 0:
                    continue
                
                # Colonne 0 = Nom/Prénom/Téléphone
                name_cell = flatten_cell_content(row[0])
                
                if not name_cell or not name_cell.strip():
                    continue
                
                # Séparer les clients multiples dans la même cellule
                client_texts = split_clients_in_cell(name_cell)
                
                # Collecter toutes les dates de séances des colonnes suivantes
                all_dates = []
                for col_idx in range(1, len(row)):
                    date_cell = flatten_cell_content(row[col_idx])
                    if date_cell:
                        dates = parse_dates(date_cell)
                        all_dates.extend(dates)
                
                # Dédupliquer les dates
                unique_dates = sorted(list(set(all_dates)))
                
                # Traiter chaque client de la cellule
                for client_text in client_texts:
                    nom, prenom, telephone = parse_name(client_text)
                    
                    # Ignorer les entrées sans nom
                    if not nom:
                        continue
                    
                    # Créer l'entrée client
                    client = {
                        'nom': nom,
                        'prenom': prenom,
                        'telephone': telephone,
                        'dates': unique_dates.copy(),
                        'nb_seances': len(unique_dates),
                        'source_file': os.path.basename(filepath)
                    }
                    
                    clients.append(client)
        
        print(f"   ✓ {len(clients)} clients extraits")
        
    except Exception as e:
        print(f"   ✗ Erreur lors du traitement: {e}")
    
    return clients


def process_all_docx_files(input_dir: str) -> List[Dict]:
    """
    Traite tous les fichiers DOCX dans le répertoire spécifié.
    
    Args:
        input_dir: Chemin vers le répertoire contenant les fichiers DOCX
        
    Returns:
        Liste complète de tous les clients extraits
    """
    all_clients = []
    
    # Vérifier que le répertoire existe
    if not os.path.exists(input_dir):
        print(f"❌ Erreur: Le répertoire '{input_dir}' n'existe pas.")
        return all_clients
    
    # Trouver tous les fichiers DOCX
    docx_files = list(Path(input_dir).glob("*.docx"))
    
    # Filtrer les fichiers temporaires de Word (commençant par ~$)
    docx_files = [f for f in docx_files if not f.name.startswith('~$')]
    
    if not docx_files:
        print(f"❌ Aucun fichier DOCX trouvé dans '{input_dir}'")
        return all_clients
    
    print(f"\n🔍 {len(docx_files)} fichier(s) DOCX trouvé(s)")
    
    # Traiter chaque fichier
    for docx_file in docx_files:
        clients = extract_clients_from_docx(str(docx_file))
        all_clients.extend(clients)
    
    print(f"\n📊 Total: {len(all_clients)} entrées clients extraites")
    
    return all_clients


# ============================================================================
# FUSION DES DOUBLONS
# ============================================================================

def calculate_similarity(client1: Dict, client2: Dict) -> float:
    """
    Calcule un score de similarité entre deux clients basé sur Nom + Prénom.
    
    Args:
        client1, client2: Dictionnaires clients
        
    Returns:
        Score de similarité (0-100)
    """
    name1 = f"{client1['nom']} {client1['prenom']}".strip().lower()
    name2 = f"{client2['nom']} {client2['prenom']}".strip().lower()
    
    # Utiliser le ratio de Levenshtein
    return fuzz.ratio(name1, name2)


def merge_duplicate_clients(clients: List[Dict], threshold: int = FUZZY_THRESHOLD) -> List[Dict]:
    """
    Fusionne les clients doublons en utilisant fuzzy matching.
    
    Args:
        clients: Liste de clients
        threshold: Score minimum pour considérer deux clients comme identiques
        
    Returns:
        Liste de clients fusionnés
    """
    if not clients:
        return []
    
    print(f"\n🔄 Fusion des doublons (seuil: {threshold}%)...")
    
    merged = []
    processed = set()
    
    for i, client1 in enumerate(clients):
        if i in processed:
            continue
        
        # Créer un nouveau client fusionné
        merged_client = {
            'nom': client1['nom'],
            'prenom': client1['prenom'],
            'telephone': client1['telephone'],
            'dates': set(client1['dates']),
            'source_files': {client1['source_file']}
        }
        
        # Chercher les doublons
        for j, client2 in enumerate(clients[i+1:], start=i+1):
            if j in processed:
                continue
            
            similarity = calculate_similarity(client1, client2)
            
            if similarity >= threshold:
                # Fusionner les données
                merged_client['dates'].update(client2['dates'])
                merged_client['source_files'].add(client2['source_file'])
                
                # Prendre le téléphone le plus complet
                if len(client2['telephone']) > len(merged_client['telephone']):
                    merged_client['telephone'] = client2['telephone']
                
                processed.add(j)
        
        processed.add(i)
        
        # Convertir les dates en liste triée
        merged_client['dates'] = sorted(list(merged_client['dates']))
        merged_client['nb_seances'] = len(merged_client['dates'])
        
        merged.append(merged_client)
    
    print(f"   ✓ {len(clients)} → {len(merged)} clients après fusion")
    
    return merged


# ============================================================================
# FILTRAGE ET EXPORT
# ============================================================================

def filter_loyal_clients(clients: List[Dict], min_sessions: int = MIN_SESSIONS) -> List[Dict]:
    """
    Filtre les clients ayant au moins le nombre minimum de séances.
    
    Args:
        clients: Liste de clients
        min_sessions: Nombre minimum de séances requis
        
    Returns:
        Liste filtrée de clients fidèles
    """
    loyal = [c for c in clients if c['nb_seances'] >= min_sessions]
    
    print(f"\n✅ {len(loyal)} clients fidèles (≥{min_sessions} séances) sur {len(clients)} au total")
    
    return loyal


def export_to_csv(clients: List[Dict], output_file: str):
    """
    Exporte la liste de clients vers un fichier CSV.
    
    Args:
        clients: Liste de clients à exporter
        output_file: Chemin du fichier CSV de sortie
    """
    if not clients:
        print("⚠️  Aucun client à exporter")
        return
    
    # Trier par nombre de séances (décroissant) puis par nom
    clients_sorted = sorted(
        clients,
        key=lambda c: (-c['nb_seances'], c['nom'], c['prenom'])
    )
    
    # Créer le répertoire de sortie si nécessaire
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Écrire le CSV
    with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
        fieldnames = ['Nom', 'Prénom', 'Téléphone', 'Nombre de séances']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        
        for client in clients_sorted:
            writer.writerow({
                'Nom': client['nom'],
                'Prénom': client['prenom'],
                'Téléphone': client['telephone'],
                'Nombre de séances': client['nb_seances']
            })
    
    print(f"\n💾 Fichier exporté: {output_file}")
    print(f"   {len(clients_sorted)} clients enregistrés")


def export_to_odoo_format(clients: List[Dict], output_file: str):
    """
    Exporte la liste de clients au format Odoo (conforme au template).
    
    Args:
        clients: Liste de clients à exporter
        output_file: Chemin du fichier CSV de sortie
    """
    if not clients:
        print("⚠️  Aucun client à exporter")
        return
    
    # Trier par nombre de séances (décroissant) puis par nom
    clients_sorted = sorted(
        clients,
        key=lambda c: (-c['nb_seances'], c['nom'], c['prenom'])
    )
    
    # Créer le répertoire de sortie si nécessaire
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Préparer les données au format Odoo
    odoo_data = []
    for client in clients_sorted:
        # Créer le nom complet (Prénom Nom)
        if client['prenom']:
            nom_complet = f"{client['prenom']} {client['nom']}"
        else:
            nom_complet = client['nom']
        
        odoo_data.append({
            'Name': nom_complet,
            'Phone': client['telephone'],
            'Tags': 'Client Fidèle',
            'Notes': f"Nombre de séances: {client['nb_seances']}"
        })
    
    # Écrire le CSV au format Odoo
    with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
        fieldnames = ['Name', 'Phone', 'Tags', 'Notes']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        writer.writerows(odoo_data)
    
    print(f"\n💾 Fichier Odoo exporté: {output_file}")
    print(f"   {len(odoo_data)} clients enregistrés")
    print(f"   ✅ Format compatible avec le template Odoo Contacts")


def display_statistics(clients: List[Dict]):
    """
    Affiche des statistiques sur les clients extraits.
    
    Args:
        clients: Liste de clients
    """
    if not clients:
        return
    
    print("\n" + "="*60)
    print("📈 STATISTIQUES")
    print("="*60)
    
    # Nombre total de clients
    print(f"Clients fidèles identifiés: {len(clients)}")
    
    # Distribution par nombre de séances
    session_counts = {}
    for client in clients:
        nb = client['nb_seances']
        session_counts[nb] = session_counts.get(nb, 0) + 1
    
    print("\nDistribution par nombre de séances:")
    for nb in sorted(session_counts.keys(), reverse=True):
        print(f"  {nb} séances: {session_counts[nb]} clients")
    
    # Top 5 clients
    top_clients = sorted(clients, key=lambda c: c['nb_seances'], reverse=True)[:5]
    print("\n🏆 Top 5 clients les plus fidèles:")
    for i, client in enumerate(top_clients, 1):
        print(f"  {i}. {client['nom']} {client['prenom']} - {client['nb_seances']} séances")
    
    # Clients avec et sans téléphone
    with_phone = sum(1 for c in clients if c['telephone'])
    print(f"\nClients avec téléphone: {with_phone}/{len(clients)} ({with_phone*100//len(clients)}%)")
    
    print("="*60)


# ============================================================================
# FONCTION PRINCIPALE
# ============================================================================

def main():
    """
    Fonction principale du script.
    """
    print("="*60)
    print("🌟 EXTRACTION DES CLIENTS FIDÈLES - RADIANCE CRM")
    print("="*60)
    
    # 1. Extraire tous les clients de tous les fichiers DOCX
    all_clients = process_all_docx_files(INPUT_DIR)
    
    if not all_clients:
        print("\n❌ Aucun client extrait. Vérifiez vos fichiers DOCX.")
        return
    
    # 2. Fusionner les doublons
    merged_clients = merge_duplicate_clients(all_clients, FUZZY_THRESHOLD)
    
    # 3. Filtrer les clients fidèles (≥ MIN_SESSIONS séances)
    loyal_clients = filter_loyal_clients(merged_clients, MIN_SESSIONS)
    
    if not loyal_clients:
        print("\n⚠️  Aucun client fidèle trouvé avec le critère minimum.")
        return
    
    # 4. Afficher les statistiques
    display_statistics(loyal_clients)
    
    # 5. Exporter vers CSV (format standard)
    export_to_csv(loyal_clients, OUTPUT_FILE)
    
    # 6. Exporter vers CSV (format Odoo)
    odoo_file = OUTPUT_FILE.replace('.csv', '_odoo.csv')
    export_to_odoo_format(loyal_clients, odoo_file)
    
    print("\n✅ Traitement terminé avec succès!")
    print(f"\n📁 Fichiers générés:")
    print(f"   • {OUTPUT_FILE}")
    print(f"   • {odoo_file} ← À importer dans Odoo")
    print(f"\n💡 Pour importer dans Odoo:")
    print(f"   1. Menu Contacts → Favoris → Importer")
    print(f"   2. Chargez: {os.path.basename(odoo_file)}")
    print(f"   3. Vérifiez le mapping et importez")


if __name__ == "__main__":
    main()
