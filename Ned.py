import pandas as pd
from wiktionaryparser import WiktionaryParser
import time

parser = WiktionaryParser()
parser.set_default_language("dutch")

df = pd.read_excel("vocabulary.xlsx")
result = parser.fetch("neus")
print(result)

"""def get_article_info(word):
    try:
        result = parser.fetch(word)
        if not result:
            return {"Article": "not found", "POS": ""}
        pos_list = []
        for definition in result[0]["definitions"]:
            pos = definition["partOfSpeech"]
            pos_list.append(pos)
            if "noun" in pos.lower():
                return {"Article": definition.get("article", "no article found"),
                        "POS": ", ".join(pos_list)}
        # No noun found but definitions exist
        return {"Article": "not a noun", "POS": ", ".join(pos_list)}
    except Exception as e:
        print(f"Error fetching {word}: {e}")
        return {"Article": "error", "POS": ""}

# Apply to first 10 rows for testing
df_test = df.head(10).copy()
article_info = [get_article_info(w) for w in df_test["Ned"]]

df_test["Article"] = [info["Article"] for info in article_info]
df_test["POS"] = [info["POS"] for info in article_info]

df_test.to_excel("vocabulary_ugraded_test.xlsx", index=False)
print("Done! Test file saved as vocabulary_ugraded_test.xlsx")
"""