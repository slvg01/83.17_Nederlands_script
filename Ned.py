import pandas as pd
from wiktionaryparser import WiktionaryParser

# Charger le fichier d'origine
file_path = r"P:\8 - Sylvain PERSO\7 - Jobsrch\10 - New Job\00 - GIT_PERSO\17_Nederlands_script\Vocabulaire.xlsx"
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()

# Initialiser le parser
parser = WiktionaryParser()
parser.set_default_language('dutch')

# Fonction pour déterminer l'article
def get_article(word):
    word = word.lower()
    if word.endswith(("je", "tje", "pje", "etje", "mpje")):
        return "het"
    if word.endswith(("isme", "sel", "um", "ment")):
        return "het"
    if any(word.startswith(prefix) for prefix in ["ge", "be", "ver", "ont"]) and len(word) > 4:
        return "het"
    if word.endswith(("ing", "ij", "er", "aar", "heid", "ie")):
        return "de"
    if word.endswith("en"):  # pluriel
        return "de"
    return "de"  # fallback

not_found = []

# Analyse des mots
parts = []
articles = []

for word in df["Nederlands"]:
    try:
        data = parser.fetch(word)
        if not data or not data[0]["definitions"]:
            parts.append(None)
            articles.append(None)
            not_found.append(word)
            continue

        pos = data[0]["partOfSpeech"]  # nature grammaticale
        parts.append(pos)

        if pos == "noun":
            articles.append(get_article(word))
        else:
            articles.append(None)
    except Exception as e:
        parts.append(None)
        articles.append(None)
        not_found.append(word)

# Ajouter les résultats dans le DataFrame
df["PartOfSpeech"] = parts
df["Article"] = articles

# Sauvegarder le fichier corrigé
df.to_excel("Classeur1_corrigé.xlsx", index=False)

# Sauvegarder la liste des mots non trouvés
with open("not_found.txt", "w", encoding="utf-8") as f:
    for w in not_found:
        f.write(str(w) + "\n")

print("✅ Fichier Classeur1_corrigé.xlsx créé.")
print("⚠️ Les mots non trouvés sont listés dans not_found.txt")
