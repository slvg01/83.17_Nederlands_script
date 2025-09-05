from playwright.sync_api import sync_playwright

def get_nouns_with_articles(word):
    url = f"https://woordenlijst.org/zoeken/?q={word}"
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url)

        # attendre que lemma_region soit chargé
        page.wait_for_selector("#lemma_region", timeout=5000)

        # récupérer tous les blocs lemma_region
        regions = page.query_selector_all("#lemma_region")
        results = []

        for region in regions:
            # vérifie s'il y a "zelfstandig naamwoord" dans lemma_head_cat
            cat_elem = region.query_selector(".lemma_head_cat")
            if cat_elem:
                cat_text = cat_elem.inner_text().strip()
                if "zelfstandig naamwoord" in cat_text.lower():
                    # article : premier <span>
                    spans = region.query_selector_all("span")
                    article = spans[0].inner_text().strip() if spans else ""
                    # mot : span avec itemprop="name"
                    name_elem = region.query_selector('span[itemprop="name"]')
                    name = name_elem.inner_text().strip() if name_elem else ""
                    results.append((name, article))
        
        browser.close()
        return results

# Test
words = ["neus", "meisje", "mooi", "lopen", "huis"]
for w in words:
    res = get_nouns_with_articles(w)
    for name, article in res:
        print(f"{name} -> {article}")
