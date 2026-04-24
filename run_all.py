import subprocess

print("Mise à jour de la base...")
subprocess.run(["python", "update_base_vie.py"])

print("\nCalcul du scoring...")
subprocess.run(["python", "scoring_vie.py"])

print("\nAlimentation de l'onglet Travail...")
subprocess.run(["python", "alimenter_travail.py"])

print("\nTerminé.")
