# Kontrakt D/s – Generator kontraktów BDSM

Prosta, elegancka aplikacja webowa do tworzenia kontraktów Dominujący/a – Uległy/a.

## Funkcje

- **Tryb „Tworzę nowy kontrakt”** – jedna osoba wypełnia swoją część (potrzeby, oczekiwania, hard/soft limits, aftercare) i generuje plik Word.
- **Tryb „Uzupełniam otrzymany plik”** – druga osoba wgrywa plik, uzupełnia swoją część i generuje wersję kompletną.
- Profesjonalny wygląd dokumentu Word (preambuła o dobrowolności i braku mocy prawnej, safewords, podpisy itd.).
- Ciemny, stonowany interfejs.
- Wszystko działa lokalnie – żadne dane nie są wysyłane na zewnętrzny serwer.

## Uruchomienie lokalne

```bash
# 1. Wejdź do folderu projektu
cd bdsm_contract_app

# 2. (opcjonalnie) utwórz środowisko wirtualne
python -m venv venv
source venv/bin/activate        # Linux/macOS
# lub
venv\Scripts\activate           # Windows

# 3. Zainstaluj zależności
pip install -r requirements.txt

# 4. Uruchom aplikację
streamlit run app.py
