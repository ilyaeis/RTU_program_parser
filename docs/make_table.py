"""Build the example tables in README.md from examples/result.xlsx (no scraping)."""
import pandas as pd

df = pd.read_excel("examples/result.xlsx")
print(f"{len(df)} extra lectures\n")

days = ["Pirmdiena", "Otrdiena", "Trešdiena", "Ceturtdiena", "Piektdiena"]
counts = (df[df["Diena"].isin(days)]
          .pivot_table(index="Diena", columns="Nedeļas kārta", values="Nosaukums", aggfunc="count")
          .reindex(days)[["Pirmā nedeļa", "Otrā nedeļa"]])
print(counts.to_markdown())
print()
print(df.head(5)[["Nedeļas kārta", "Diena", "Laiks", "Nosaukums", "Grupu skaits"]].to_markdown(index=False))
