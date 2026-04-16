import pandas as pd

FILE = "base_offres_vie.xlsx"


def clean_text(text):
    if pd.isna(text):
        return ""
    return str(text).lower()


def is_valid_date(date_str):
    if pd.isna(date_str):
        return False

    try:
        date = pd.to_datetime(date_str)
        return date >= pd.Timestamp("2026-09-01")
    except:
        return False


def score_location(text):
    text = clean_text(text)
    score = 0

    if any(x in text for x in ["seoul", "corée", "korea"]):
        score += 5

    if any(x in text for x in ["tokyo", "japan"]):
        score += 5

    if "singapore" in text:
        score += 5

    if any(x in text for x in ["china", "shanghai", "hong kong", "taiwan"]):
        score += 4

    if any(x in text for x in ["usa", "new york", "san francisco", "boston"]):
        score += 3

    return score


def score_profile(text):
    text = clean_text(text)
    score = 0

    strong = [
        "analyst", "data", "business", "operations",
        "project", "strategy", "finance", "performance",
        "reporting", "dashboard", "pilotage"
    ]

    medium = [
        "coordination", "process", "crm", "sales",
        "consulting", "gestion", "planning"
    ]

    negative = [
        "rh", "human resources", "legal", "juridique",
        "pharma", "laboratory", "technician", "nurse",
        "mechanic", "maintenance"
    ]

    for word in strong:
        if word in text:
            score += 2

    for word in medium:
        if word in text:
            score += 1

    for word in negative:
        if word in text:
            score -= 3

    return score


def classify(score):
    if score < 0:
        return "Hors cible"
    if score >= 8:
        return "Top priorité"
    if score >= 5:
        return "Très intéressant"
    if score >= 3:
        return "À regarder"
    return "Secondaire"


def main():
    print("Chargement du fichier...")

    df = pd.read_excel(FILE, sheet_name="Base")

    scores_loc = []
    scores_prof = []
    total_scores = []
    priorities = []
    valid_flags = []

    for _, row in df.iterrows():
        poste = clean_text(row.get("poste"))
        ville = clean_text(row.get("ville"))
        desc = clean_text(row.get("description_mission"))
        date = row.get("date_debut")

        full_text = f"{poste} {ville} {desc}"

        valid = is_valid_date(date)

        if not valid:
            scores_loc.append(0)
            scores_prof.append(0)
            total_scores.append(-999)
            priorities.append("Hors cible (date)")
            valid_flags.append(False)
            continue

        loc_score = score_location(full_text)
        prof_score = score_profile(full_text)
        total = loc_score + prof_score

        scores_loc.append(loc_score)
        scores_prof.append(prof_score)
        total_scores.append(total)
        priorities.append(classify(total))
        valid_flags.append(True)

    df["score_localisation"] = scores_loc
    df["score_profil"] = scores_prof
    df["score_total"] = total_scores
    df["priorite"] = priorities
    df["valide_date"] = valid_flags

    df_valid = df[df["valide_date"] == True].copy()
    df_valid = df_valid.sort_values(by="score_total", ascending=False)

    top50 = df_valid.head(50)

    print("Écriture dans le fichier...")

    with pd.ExcelWriter(FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Base", index=False)
        top50.to_excel(writer, sheet_name="Top priorite", index=False)

    print("Terminé. Onglet 'Top priorite' créé.")


if __name__ == "__main__":
    main()