import pandas as pd
from playwright.sync_api import sync_playwright
import time


def fetch_all_categories(word, page):
    """Scrape toutes les catégories pour un mot donné (format uniforme)."""
    url = f"https://woordenlijst.org/zoeken/?q={word}"
    page.goto(url)

    try:
        page.wait_for_selector("#lemma_region", timeout=5000)
    except:
        return [("", word, "")]  # si rien trouvé → juste le mot du fichier

    results = []
    regions = page.query_selector_all("#lemma_region")

    for region in regions:
        categorie_elem = region.query_selector(".lemma_head_cat")
        categorie = categorie_elem.inner_text().strip() if categorie_elem else "not found"

        name_elem = region.query_selector('span[itemprop="name"]')
        name_found = name_elem.inner_text().strip() if name_elem else "missing"

        article = ""
        if "zelfstandig naamwoord" in categorie.lower():
            spans = region.query_selector_all("span")
            article = spans[0].inner_text().strip() if spans else ""

        results.append((article, name_found, categorie))

    if not results:
        return [("", word, "")]

    return results


def build_vocabulary(input_file="Vocabulary_new.xlsx", output_file="Vocabulary_new_upgraded.xlsx"):
    """Construit un Excel enrichi avec les catégories trouvées sur woordenlijst.org"""
    df = pd.read_excel(input_file)

    all_results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for word in df["Nederlands"].tolist():
            site_results = fetch_all_categories(word, page)

            # supprimer les doublons avec le mot original
            filtered_results = [r for r in site_results if r[1].lower() != word.lower()]

            # ordre final = mot du fichier en premier + résultats du site
            final_row = [("", word, "")] + filtered_results
            all_results.append(final_row)

            time.sleep(0.5)  # petit délai pour ne pas surcharger le site

        browser.close()

    # déterminer le max de colonnes nécessaires
    max_len = max(len(r) for r in all_results)

    # construire les colonnes
    columns = {}
    for i in range(max_len):
        columns[f"Article_{i+1}"] = []
        columns[f"Name_{i+1}"] = []
        columns[f"Category_{i+1}"] = []

    # remplir les colonnes
    for row in all_results:
        for i in range(max_len):
            if i < len(row):
                a, n, c = row[i]
            else:
                a, n, c = "", "", ""
            columns[f"Article_{i+1}"].append(a)
            columns[f"Name_{i+1}"].append(n)
            columns[f"Category_{i+1}"].append(c)

    # ajout au DataFrame original
    for col, values in columns.items():
        df[col] = values

    # sauvegarde
    df.to_excel(output_file, index=False)
    print(f"✅ Updated Excel saved as {output_file}")


# --- utilisation ---
if __name__ == "__main__":
    build_vocabulary()
