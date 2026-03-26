import pandas as pd

try:
    df = pd.read_excel('PLANIFICACIÓN ANUAL 2026(1-63).xlsx')
    with open('cols.txt', 'w', encoding='utf-8') as f:
        for c in df.columns:
            f.write(repr(c) + '\n')
    print("Columns written to cols.txt")
except Exception as e:
    print("Error:", e)
