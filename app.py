import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import openpyxl

# Konfiguracja strony (musi być pierwszą komendą)
st.set_page_config(
    page_title="URL Matcher",
    page_icon="🚀",
    layout="wide"
)

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
        # Optymalizacja: List comprehension jest szybkie, ale w Streamlit chcemy to zrobić raz
        
        # 1. Kandydaci z kolumny B (URL2)
        all_url2_values = [
            str(cell.value).strip() 
            for cell in sheet['B'][1:] 
            if cell.value is not None and str(cell.value).strip() != ''
        ]
        candidates = list(set(all_url2_values)) # Unikalne wartości
        
        # 2. Cele z kolumny A (URL1) i odpowiadające im URL2 (dla kontekstu)
        url1_data = []
        for cell_a in sheet['A'][1:]:
            if cell_a.value is not None and str(cell_a.value).strip() != '':
                # Pobranie wartości z kolumny B w tym samym wierszu
                url2_val = sheet.cell(row=cell_a.row, column=2).value
                url2_str = str(url2_val) if url2_val is not None else ''
                
                url1_data.append({
                    'URL1': str(cell_a.value).strip(),
                    'URL2': url2_str
                })
        
        if not url1_data:
            return None, "Nie znaleziono danych w kolumnie URL1 (kolumna A)."
            
        return url1_data, candidates

    except Exception as e:
        return None, f"Błąd podczas odczytu pliku: {e}"

def run_matching(url1_data, candidates):
    results = []
    total_items = len(url1_data)
    
    # Pasek postępu
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, item in enumerate(url1_data):
        url1 = item['URL1']
        
        # Logika RapidFuzz
        # limit=3 zwraca listę krotek: (match, score, index)
        matches = process.extract(url1, candidates, scorer=fuzz.WRatio, limit=3)
        
        row = {
            'URL1': url1,
            'URL2 (Oryginał)': item['URL2']
        }
        
        for i, match in enumerate(matches):
            url_match, similarity, _ = match
            row[f'Propozycja {i+1}'] = url_match
            row[f'% {i+1}'] = f"{similarity:.2f}"
            
        results.append(row)
        
        # Aktualizacja paska postępu co 10 lub co 1% (dla wydajności)
        if idx % 10 == 0 or idx == total_items - 1:
            progress = (idx + 1) / total_items
            progress_bar.progress(progress)
            status_text.text(f"Przetwarzanie: {idx + 1}/{total_items}")

    progress_bar.empty()
    status_text.empty()
    
    return pd.DataFrame(results)

# --- Interfejs Użytkownika (UI) ---

st.title("🚀 URL Matcher Pro")
st.markdown("""
To narzędzie dopasowuje adresy URL z **Kolumny A (URL1)** do najbardziej podobnych adresów z **Kolumny B (URL2)**.
Działa nawet, jeśli dane w Excelu są wynikami formuł.
""")

st.info("Wymagany format pliku: Excel (.xlsx), dane w pierwszej zakładce. Kolumna A: URL do sprawdzenia, Kolumna B: Baza adresów.")

uploaded_file = st.file_uploader("Wgraj plik Excel (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    # Przycisk startu, żeby użytkownik miał kontrolę
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
