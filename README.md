# redirections
🚀 Aplikacja webowa w Streamlit do inteligentnego dopasowywania i analizy podobieństwa adresów URL (Fuzzy Matching) z plików Excel. Szybka, obsługuje formuły i generuje raporty TOP 3.

# 🚀 URL Matcher Pro

Narzędzie do automatycznego dopasowywania i analizy podobieństwa adresów URL (Fuzzy Matching), stworzone w Pythonie przy użyciu frameworka **Streamlit**.

Aplikacja pozwala na wgranie pliku Excel z dwiema kolumnami adresów, a następnie dla każdego adresu z pierwszej kolumny znajduje 3 najbardziej zbliżone odpowiedniki w drugiej kolumnie.

🔗 **[Uruchom aplikację na Streamlit Cloud](https://share.streamlit.io/)**

## ✨ Główne funkcjonalności

*   **Inteligentne dopasowanie (Fuzzy Matching):** Wykorzystuje bibliotekę `RapidFuzz` (algorytm WRatio) do znajdowania podobieństw, nawet przy literówkach czy drobnych różnicach w strukturze URL.
*   **Obsługa formuł Excela:** Dzięki trybowi `data_only=True`, skrypt poprawnie odczytuje wartości wygenerowane przez formuły (np. VLOOKUP czy CONCATENATE), a nie sam kod formuły.
*   **Wyniki TOP 3:** Dla każdego adresu zwraca trzy najlepsze propozycje wraz z procentowym wskaźnikiem podobieństwa.
*   **Szybkość i wydajność:** Zoptymalizowane działanie dzięki cache'owaniu wyników – zmiana parametrów wyświetlania nie wymaga ponownego przeliczania danych.
*   **Prosty interfejs:** Wgrywasz plik -> Klikasz "Start" -> Pobierasz gotowy raport.

## 📂 Wymagany format danych

Aplikacja oczekuje pliku **Excel (.xlsx)** przygotowanego w następujący sposób:

1.  **Arkusz:** Dane muszą znajdować się w pierwszym (aktywnym) arkuszu.
2.  **Kolumny:**
    *   **Kolumna A (URL1):** Lista adresów, dla których szukamy dopasowań (to co sprawdzamy).
    *   **Kolumna B (URL2):** Baza adresów (wzorzec), w której szukamy kandydatów.
3.  **Nagłówki:** Wiersz 1 jest traktowany jako nagłówek i jest pomijany w analizie.

## 🛠️ Instalacja i uruchomienie lokalne

Jeśli chcesz uruchomić aplikację na własnym komputerze zamiast w chmurze:

1.  **Sklonuj repozytorium:**
    ```bash
    git clone https://github.com/TWOJA_NAZWA_UZYTKOWNIKA/url-matcher.git
    cd url-matcher
    ```

2.  **Zainstaluj wymagane biblioteki:**
    Zalecane jest użycie wirtualnego środowiska (venv).
    ```bash
    pip install -r requirements.txt
    ```

3.  **Uruchom aplikację:**
    ```bash
    streamlit run app.py
    ```
    Aplikacja otworzy się automatycznie w Twojej domyślnej przeglądarce pod adresem `http://localhost:8501`.

## 📦 Wykorzystane technologie

*   [Streamlit](https://streamlit.io/) - Interfejs użytkownika (Web App).
*   [Pandas](https://pandas.pydata.org/) - Przetwarzanie danych i obsługa ramek danych.
*   [RapidFuzz](https://github.com/maxbachmann/RapidFuzz) - Bardzo szybka biblioteka do dopasowywania ciągów znaków (String Matching).
*   [OpenPyXL](https://openpyxl.readthedocs.io/) - Odczyt i zapis plików Excel (.xlsx).

---
