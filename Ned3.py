import pandas as pd
from playwright.sync_api import sync_playwright
import time

df = pd.read_excel("vocabulary.xlsx")

def fetch_all_categories(word, page):
    url = f"https://woordenlijst.org/zoeken/?q={word}"
    page.goto(url)
    try:
        page.wait_for_selector("#lemma_region", timeout=5000)
    except:
        return [("", "", "")]  # Article, Name, Category for first, then empty tuples

    nouns = []
    others = []
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
            nouns.append((article, name_found, categorie))
        else:
            others.append((name_found, categorie))

    # Return list: first tuple is (Article, Name, Category), rest are (Name, Category)
    combined = nouns + others
    if not combined:
        return [("", "", "")]
    return combined

# Fetch results
all_results = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    for word in df["Ned"].tolist():
        formatted = fetch_all_categories(word, page)
        all_results.append(formatted)
        time.sleep(0.2)

    browser.close()

# Determine max number of categories
max_len = max(len(r) for r in all_results)

# Prepare columns
columns = {}
columns["Article_1"] = []
columns["Name_1"] = []
columns["Category_1"] = []

for i in range(1, max_len):
    columns[f"Name_{i+1}"] = []
    columns[f"Category_{i+1}"] = []

# Fill columns
for row in all_results:
    # First category
    first = row[0] if row else ("", "", "")
    if len(first) == 3:
        article, name, cat = first
    else:  # it's not a noun
        article = ""
        name, cat = first
    columns["Article_1"].append(article)
    columns["Name_1"].append(name)
    columns["Category_1"].append(cat)

    # Other categories
    for i in range(1, max_len):
        if i < len(row):
            if len(row[i]) == 2:
                n, c = row[i]
                columns[f"Name_{i+1}"].append(n)
                columns[f"Category_{i+1}"].append(c)
            else:  # fallback if somehow tuple has 3
                _, n, c = row[i]
                columns[f"Name_{i+1}"].append(n)
                columns[f"Category_{i+1}"].append(c)
        else:
            columns[f"Name_{i+1}"].append("")
            columns[f"Category_{i+1}"].append("")

# Add to dataframe
for col, values in columns.items():
    df[col] = values

# Save Excel
df.to_excel("vocabulary_upgraded.xlsx", index=False)
print("Updated Excel saved as vocabulary_upgraded_first_article.xlsx")
