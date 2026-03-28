import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import openpyxl

# 1. Konfiguracja strony (musi być ZAWSZE pierwsza)
st.set_page_config(
    page_title="URL Matcher",
    page_icon="🔒",
    layout="wide"
)

# --- MODUŁ LOGOWANIA ---
def check_password():
    """Zwraca `True` jeśli użytkownik podał poprawne hasło."""

    def password_entered():
        """Sprawdza czy wpisane hasło zgadza się z tym w sekretach."""
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Nie przechowujemy hasła
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Pierwsze uruchomienie, pokaż formularz
        st.text_input(
            "Podaj hasło dostępu:", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.info("Dane potrzebne do zalogowania znajdują się w Monday. Kontakt: jaroslaw.muzyka@performance-group.pl")
        return False
    elif not st.session_state["password_correct"]:
        # Hasło błędne
        st.text_input(
            "Podaj hasło dostępu:", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("😕 Niepoprawne hasło")
        st.info("Dane potrzebne do zalogowania znajdują się w Monday. Kontakt: jaroslaw.muzyka@performance-group.pl")
        return False
    else:
        # Hasło poprawne
        return True

if not check_password():
    st.stop()  # Zatrzymuje aplikację, jeśli brak autoryzacji

# =========================================================
# WŁAŚCIWA APLIKACJA (Kod wykonuje się tylko po zalogowaniu)
# =========================================================

# --- Funkcja przetwarzająca (zoptymalizowana i cache'owana) ---
@st.cache_data(show_spinner=False)
def process_excel(file_bytes):
    """
    Funkcja przetwarza plik Excel w pamięci, zachowując logikę data_only=True
    dla poprawnych wartości z formuł.
    """
    try:
        # Wczytanie skoroszytu z data_only=True (kluczowe dla formuł)
        workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet = workbook.active
        
        # Pobieranie danych (pomijamy nagłówek - row 1)
        # 1. Kandydaci z kolumny B (URL2)
        all_url2_values = [
            str(cell.value).strip() 
            for cell in sheet['B'][1:] 
            if cell.value is not None and str(cell.value).strip() != ''
        ]
        candidates = list(set(all_url2_values)) # Unikalne wartości
        
        # 2. Cele z kolumny A (URL1)
        url1_data = []
        for cell_a in sheet['A'][1:]:
            if cell_a.value is not None and str(cell_a.value).strip() != '':
                url1_data.append({
                    'URL1': str(cell_a.value).strip()
                })
        
        if not url1_data:
            return None, "Nie znaleziono danych w kolumnie URL1 (kolumna A)."
            
        return url1_data, candidates

    except Exception as e:
        return None, f"Błąd podczas odczytu pliku: {e}"

def run_matching(url1_data, candidates):
    results = []
    total_items = len(url1_data)
    
    stop_btn_placeholder = st.empty()
    stop_btn_placeholder.button("🛑 Zatrzymaj analizę", key="stop_analysis_btn")

    # Pasek postępu
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, item in enumerate(url1_data):
        url1 = item['URL1']
        
        # Logika RapidFuzz
        # limit=3 zwraca listę krotek: (match, score, index)
        matches = process.extract(url1, candidates, scorer=fuzz.WRatio, limit=3)
        
        row = {
            'URL 1 (do przekierowania)': url1
        }
        
        for i, match in enumerate(matches):
            url_match, similarity, _ = match
            row[f'Propozycja {i+1}'] = url_match
            row[f'% {i+1}'] = similarity
            
        results.append(row)
        
        # Aktualizacja paska postępu co 10 lub co 1% (dla wydajności)
        if idx % 10 == 0 or idx == total_items - 1:
            progress = (idx + 1) / total_items
            progress_bar.progress(progress)
            status_text.text(f"Przetwarzanie: {idx + 1}/{total_items}")

    progress_bar.empty()
    status_text.empty()
    stop_btn_placeholder.empty()
    
    df = pd.DataFrame(results)
    
    if '% 1' in df.columns:
        df = df.sort_values(by='% 1', ascending=False)
        
    for i in range(1, 4):
        col_name = f'% {i}'
        if col_name in df.columns:
            df[col_name] = df[col_name].apply(lambda x: f"{x:.2f}")
            
    return df

# --- Interfejs Użytkownika (UI) ---

st.title("🚀 URL Matcher")
st.markdown("""
Narzędzie rozwiązuje problem masowego mapowania przekierowań 301, znajdując najbardziej zbliżone adresy URL z aktualnej oferty dla starych lub wygasłych linków.
""")

st.markdown("""
### Instrukcje:
Wgraj plik XLSX z dwiema kolumnami:
* **Kolumna A:** lista adresów do przekierowania (np. adresy z kodem 404)
* **Kolumna B:** lista wszystkich aktualnych adresów URL (kandydaci do przekierowań)

Uruchom skrypt analizy podobieństwa i poczekaj na wyniki analizy.  
Pobierz wynik z propozycjami dopasowań. 

**Przykładowy plik:**
""")

try:
    with open("przykladowy-plik.xlsx", "rb") as f:
        przykladowy_plik_bytes = f.read()
    st.download_button(
        label="📥 Pobierz przykładowy plik (przykladowy-plik.xlsx)",
        data=przykladowy_plik_bytes,
        file_name="przykladowy-plik.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except FileNotFoundError:
    st.warning("Przykładowy plik (przykladowy-plik.xlsx) nie został znaleziony w głównym katalogu.")

st.info("Wymagany format pliku: Excel (.xlsx). Kolumna A: URL do sprawdzenia, Kolumna B: Baza adresów.")

uploaded_file = st.file_uploader("Wgraj plik Excel (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    if st.button("Uruchom analizę", type="primary"):
        
        with st.spinner("Wczytywanie i analiza pliku..."):
            # Krok 1: Wczytanie danych
            bytes_data = uploaded_file.getvalue()
            data_extracted, candidates_or_error = process_excel(bytes_data)
            
            if data_extracted is None:
                st.error(f"Wystąpił błąd: {candidates_or_error}")
            else:
                st.success(f"Wczytano poprawnie! Kandydatów do dopasowania: {len(candidates_or_error)}. Adresów do sprawdzenia: {len(data_extracted)}.")
                
                # Krok 2: Uruchomienie dopasowywania
                final_df = run_matching(data_extracted, candidates_or_error)
                
                # Krok 3: Wyświetlenie wyników
                st.subheader("Podgląd wyników")
                st.dataframe(final_df.head(50), use_container_width=True)
                
                # Krok 4: Eksport do Excela
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Wyniki')
                
                st.download_button(
                    label="📥 Pobierz wyniki (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"wyniki_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
