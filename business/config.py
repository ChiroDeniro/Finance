from pathlib import Path

INVOICE_DIRS = {
    "2025": r"C:\Users\chris\Documents\Professioneel\01 Optimal Data Solutions\01 Administratie\Facturen Ontvangen\2025",
    "2026": r"C:\Users\chris\Documents\Professioneel\01 Optimal Data Solutions\01 Administratie\Facturen Ontvangen\2026",
}

SUPPORTED_EXTENSIONS = [".jpg", ".jpeg", ".png", ".webp"]

OUTPUT_DIR = Path(__file__).parent.parent / "bookkeeping" / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

OLLAMA_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "llava"
OLLAMA_TIMEOUT = 60
