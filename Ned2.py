import pandas as pd
from playwright.sync_api import sync_playwright
import time

# Charger le fichier Excel
df = pd.read_excel("vocabulary.xlsx").head(10)  # seulement 10 mots pour test

def fetch_info(word, page):
    url = f"https://woordenlijst.org/zoeken/?q={word}"
    page.goto(url)
    try:
        page.wait_for_selector("#lemma_region", timeout=5000)
    except:
        return "not found", "", "missing"

    regions = page.query_selector_all("#lemma_region")
    for region in regions:
        categorie_elem = region.query_selector(".lemma_head_cat")
        categorie = categorie_elem.inner_text().strip() if categorie_elem else "not found"

        article = ""
        if "zelfstandig naamwoord" in categorie.lower():
            spans = region.query_selector_all("span")
            article = spans[0].inner_text().strip() if spans else ""

        name_elem = region.query_selector('span[itemprop="name"]')
        name_found = name_elem.inner_text().strip() if name_elem else "missing"

        if "zelfstandig naamwoord" not in categorie.lower():
            article = ""
            categorie = "not a name"

        return categorie, article, name_found

    return "not found", "", "missing"


# Utiliser Playwright
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    categories = []
    articles = []
    names_found = []

    for word in df["Ned"].tolist():
        cat, art, name = fetch_info(word, page)
        categories.append(cat)
        articles.append(art)
        names_found.append(name)
        time.sleep(0.2)  # pause pour test

    browser.close()

df["Categorie"] = categories
df["Article"] = articles
df["Mot_trouvé"] = names_found

df.to_excel("vocabulary_upgraded_test.xlsx", index=False)
print("Fichier vocabulary_upgraded_test.xlsx créé avec succès !")
