# scan_ollama.py — Facturen verwerken via Ollama llava

Verwerkt lokale factuur-afbeeldingen (JPG/PNG) via het llava vision model en schrijft de resultaten naar Excel.

## Vereisten

```
pip install requests openpyxl
```

Ollama moet draaien op `http://localhost:11434` met het llava model:

```
ollama pull llava
```

## Gebruik

```bash
cd business

python scan_ollama.py          # verwerk 2026 (default)
python scan_ollama.py 2025     # verwerk 2025
python scan_ollama.py all      # verwerk alle jaren
```

Output: `business/output/facturen_YYYY.xlsx`

## Beperkingen

- **PDFs worden overgeslagen** — llava kan alleen afbeeldingen lezen, geen PDF bestanden.
  Converteer PDFs eerst naar JPG/PNG als je ze wilt verwerken.
- Verwerking duurt ~10–60 seconden per afbeelding afhankelijk van je hardware.
