import pandas as pd
from playwright.sync_api import sync_playwright
import time, random


def fetch_all_categories(word, page):
    """Scrape toutes les catégories pour un mot donné (format uniforme)."""
    url = f"https://woordenlijst.org/zoeken/?q={word}"
    try:
        page.goto(url)
        page.wait_for_selector("#lemma_region", timeout=5000)
    except:
        return []

    results = []
    regions = page.query_selector_all("#lemma_region")

    for region in regions:
        # catégorie grammaticale
        categorie_elem = region.query_selector(".lemma_head_cat")
        categorie = categorie_elem.inner_text().strip() if categorie_elem else "not found"

        # nom trouvé
        name_elem = region.query_selector('span[itemprop="name"]')
        name_found = name_elem.inner_text().strip() if name_elem else "missing"

        # article seulement si nom
        article = ""
        if "zelfstandig naamwoord" in categorie.lower():
            spans = region.query_selector_all("span")
            article = spans[0].inner_text().strip() if spans else ""

        # 🔥 format uniforme
        results.append((article, name_found, categorie))

    return results


def build_vocabulary(input_file="vocabulary_new.xlsx", output_file="vocabulary_mined.xlsx"):
    """Construit un Excel enrichi avec les catégories trouvées (20 premières lignes seulement)."""
    df = pd.read_excel(input_file)

    all_results = []
    found_flags = []  # nouvelle colonne

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        for word in df["Nederlands"].tolist():
            site_results = fetch_all_categories(word, page)

            # chercher si le mot Excel est dans les résultats
            match = None
            for r in site_results:
                if r[1].lower() == word.lower():
                    match = r
                    break

            if match:
                # mot trouvé sur le site
                final_row = [match] + [r for r in site_results if r != match]
                found_flags.append("Yes")
            else:
                # mot pas trouvé → mettre en premier avec article/category = "not found"
                final_row = [("", word, "not found")] + site_results
                found_flags.append("No")

            all_results.append(final_row)

            # délai random pour éviter d'être bloqué
            time.sleep(random.uniform(0.5, 1.5))

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

    df["Found_in_site"] = found_flags  # ✅ nouvelle colonne

    # sauvegarde
    df.to_excel(output_file, index=False)
    print(f"✅ Test Excel saved as {output_file}")


# --- test sur 20 lignes ---
if __name__ == "__main__":
    build_vocabulary()
