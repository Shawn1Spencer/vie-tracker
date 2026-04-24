import requests
import pandas as pd
import time
import os
from datetime import datetime

SEARCH_URL = "https://civiweb-api-prd.azurewebsites.net/api/Offers/search"
DETAIL_URL_BASE = "https://civiweb-api-prd.azurewebsites.net/api/Offers/details"

HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0"
}

BASE_PAYLOAD = {
    "limit": 100,
    "skip": 0,
    "latest": ["true"],
    "activitySectorId": [],
    "companiesSizes": [],
    "countriesIds": [],
    "entreprisesIds": [0],
    "geographicZones": [],
    "missionStartDate": None,
    "missionsDurations": [],
    "missionsTypesIds": [],
    "query": None,
    "specializationsIds": [],
    "studiesLevelId": []
}

BASE_FILE = "base_offres_vie.xlsx"
MAX_NEW_OFFERS = 2000


def fetch_search_page(payload):
    response = requests.post(SEARCH_URL, headers=HEADERS, json=payload, timeout=30)
    response.raise_for_status()
    return response.json()


def fetch_offer_detail(offer_id):
    url = f"{DETAIL_URL_BASE}/{offer_id}"
    response = requests.get(url, headers=HEADERS, timeout=30)
    response.raise_for_status()
    return response.json()


def extract_offers_list(data):
    if isinstance(data, dict) and "result" in data and isinstance(data["result"], list):
        return data["result"]
    return []


def fetch_all_offer_ids():
    all_ids = []
    seen_ids = set()

    limit = BASE_PAYLOAD["limit"]
    skip = 0
    page_number = 1

    while True:
        payload = BASE_PAYLOAD.copy()
        payload["skip"] = skip

        print(f"Recherche page {page_number}, skip={skip}...")
        data = fetch_search_page(payload)
        offers = extract_offers_list(data)

        print(f"{len(offers)} offres trouvées sur cette page")

        if not offers:
            break

        for offer in offers:
            offer_id = offer.get("id")
            if offer_id and offer_id not in seen_ids:
                seen_ids.add(offer_id)
                all_ids.append(offer_id)

        if len(offers) < limit:
            break

        skip += limit
        page_number += 1
        time.sleep(0.2)

    return all_ids


def clean_datetime_for_excel(value):
    if value is None or value == "":
        return None

    try:
        dt = pd.to_datetime(value, errors="coerce", utc=True)
        if pd.isna(dt):
            return str(value)

        dt = dt.tz_localize(None)

        if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
            return dt.strftime("%Y-%m-%d")

        return dt.strftime("%Y-%m-%d %H:%M:%S")

    except Exception:
        return str(value)


def detail_to_row(detail, today_str):
    offer_id = detail.get("id")

    return {
        "id": offer_id,
        "reference": detail.get("reference"),
        "entreprise": detail.get("organizationName"),
        "poste": detail.get("missionTitle"),
        "ville": detail.get("cityName"),
        "pays_code": detail.get("countryId"),
        "pays": detail.get("countryName"),
        "type_mission": detail.get("missionType"),
        "duree_mois": detail.get("missionDuration"),
        "date_creation": clean_datetime_for_excel(detail.get("creationDate")),
        "date_debut": clean_datetime_for_excel(detail.get("missionStartDate")),
        "date_fin": clean_datetime_for_excel(detail.get("missionEndDate")),
        "indemnite": detail.get("indemnite"),
        "teletravail": detail.get("teleworkingAvailable"),
        "portee_sociale": detail.get("socialReach"),
        "contact_nom": detail.get("contactName"),
        "contact_email": detail.get("contactEmail"),
        "description_entreprise": detail.get("organizationPresentation"),
        "description_mission": detail.get("missionDescription"),
        "profil_recherche": detail.get("missionProfile"),
        "date_premiere_detection": today_str,
        "date_derniere_detection": today_str,
        "nouvelle_offre": True,
        "active_ce_jour": True,
        "lien_offre": f"https://mon-vie-via.businessfrance.fr/offres/{offer_id}" if offer_id else None
    }


def load_existing_base():
    if os.path.exists(BASE_FILE):
        print("Base existante trouvée, chargement...")
        try:
            df = pd.read_excel(BASE_FILE, engine="openpyxl")
            return df
        except Exception as e:
            print(f"Impossible de lire la base existante : {e}")
            print("On repart sur une nouvelle base vide.")
            return pd.DataFrame()
    else:
        print("Aucune base existante, création d'une nouvelle base.")
        return pd.DataFrame()


def save_base(df):
    df.to_excel(BASE_FILE, index=False, engine="openpyxl")
    print(f"Base enregistrée : {BASE_FILE}")


def ensure_update_columns(df):
    needed_defaults = {
        "date_premiere_detection": None,
        "date_derniere_detection": None,
        "nouvelle_offre": False,
        "active_ce_jour": False,
    }

    for col, default_value in needed_defaults.items():
        if col not in df.columns:
            df[col] = default_value

    # force ces colonnes en texte / objet pour éviter les conflits de type
    df["date_premiere_detection"] = df["date_premiere_detection"].astype("object")
    df["date_derniere_detection"] = df["date_derniere_detection"].astype("object")
    df["nouvelle_offre"] = df["nouvelle_offre"].astype("object")
    df["active_ce_jour"] = df["active_ce_jour"].astype("object")

    return df


def update_base():
    today_str = datetime.today().strftime("%Y-%m-%d")

    existing_df = load_existing_base()

    if not existing_df.empty:
        existing_df = ensure_update_columns(existing_df)

    if not existing_df.empty and "id" in existing_df.columns:
        existing_ids = set(pd.to_numeric(existing_df["id"], errors="coerce").dropna().astype(int).tolist())
    else:
        existing_ids = set()

    current_ids = fetch_all_offer_ids()
    current_ids_set = set(current_ids)

    print(f"\nTotal d'offres actuellement visibles : {len(current_ids)}")

    if not existing_df.empty:
        existing_df["active_ce_jour"] = False
        existing_df["nouvelle_offre"] = False

    updated_count = 0

    if not existing_df.empty and "id" in existing_df.columns:
        existing_df["id"] = pd.to_numeric(existing_df["id"], errors="coerce")

        for idx, row in existing_df.iterrows():
            offer_id = row.get("id")
            if pd.notna(offer_id) and int(offer_id) in current_ids_set:
                existing_df.at[idx, "date_derniere_detection"] = today_str
                existing_df.at[idx, "active_ce_jour"] = True
                existing_df.at[idx, "nouvelle_offre"] = False
                updated_count += 1

    new_ids = [offer_id for offer_id in current_ids if offer_id not in existing_ids]
    print(f"Nouvelles offres à enrichir : {len(new_ids)}")

    new_rows = []

    for i, offer_id in enumerate(new_ids, start=1):
        if i > MAX_NEW_OFFERS:
            print(f"Limite de test atteinte ({MAX_NEW_OFFERS}), arrêt.")
            break

        print(f"Détail {i}/{len(new_ids)} - offre {offer_id}")

        try:
            detail = fetch_offer_detail(offer_id)
            row = detail_to_row(detail, today_str)
            new_rows.append(row)
            time.sleep(0.15)

        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code == 404:
                print(f"Offre {offer_id} introuvable (404), ignorée.")
            else:
                print(f"Erreur HTTP sur {offer_id} : {e}")

        except KeyboardInterrupt:
            print("\nArrêt manuel détecté, sauvegarde de ce qui a déjà été récupéré...")
            break

        except Exception as e:
            print(f"Erreur inattendue sur {offer_id} : {e}")

    new_df = pd.DataFrame(new_rows)

    if existing_df.empty:
        final_df = new_df.copy()
    else:
        final_df = pd.concat([existing_df, new_df], ignore_index=True)

    if not final_df.empty and "id" in final_df.columns:
        final_df = final_df.drop_duplicates(subset=["id"], keep="first").reset_index(drop=True)

    save_base(final_df)

    print("\nRésumé :")
    print(f"- Offres déjà connues revues aujourd'hui : {updated_count}")
    print(f"- Nouvelles offres ajoutées : {len(new_rows)}")
    print(f"- Total dans la base : {len(final_df)}")


def main():
    print("Mise à jour de la base VIE...")

    try:
        update_base()
    except KeyboardInterrupt:
        print("\nScript arrêté manuellement.")


if __name__ == "__main__":
    main()
