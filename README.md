# Pjoter Development | Web App do analizy zamowien i wysylek

Nowoczesna aplikacja webowa w `Streamlit`, ktora porownuje poprzedni i aktualny release klienta, pokazuje wzrosty, spadki, alerty oraz generuje czytelny raport Excel.

## Co potrafi aplikacja

- porownuje dwa pliki Excel
- analizuje zmiany po `Ship Date` i `Receipt Date`
- ma kalendarz zakresu dat, ktory odswieza dashboard, tabele, wykresy i eksport
- wylicza tygodnie ISO (`YYYY-Www`) oraz ostatni pelny zakonczony tydzien referencyjny
- pokazuje agregacje tygodniowe, porownanie poprzedni release vs aktualny release i zmiane tydzien do tygodnia
- liczy dni robocze tylko wedlug kalendarza polskiego: bez sobot, niedziel i polskich swiat ustawowych
- pokazuje `Previous Qty`, `Current Qty`, `Delta` i `Percent Change`
- oznacza alerty dla zmian przekraczajacych prog `15%`
- prezentuje dane jako premium dashboard z KPI, wykresami, tabelami i macierza
- pozwala pobrac filtrowane dane jako CSV
- generuje biznesowy raport Excel z dodatkowymi arkuszami `Weekly Summary` i `Calendar PL`
- posiada ekran logowania z lokalna konfiguracja uzytkownikow
- moze zostac spakowana do uruchamianej paczki `.exe`

## Logowanie

Aplikacja ma wbudowany ekran logowania.

Plik uzytkownikow:

```text
config/users.json
```

Domyslne konto startowe:

- login: `admin`
- haslo: `Pjoter2026!`

Po pierwszym uruchomieniu zmien te dane w `config/users.json`.
W wersji EXE najwygodniej edytowac:

```text
dist\ShipmentAnalyzerWeb\config\users.json
```

Hasla sa przechowywane jako hash `PBKDF2-SHA256`.

## Jak uruchomic lokalnie

Najprostsza opcja:

1. skopiuj caly folder projektu na komputer docelowy
2. kliknij dwukrotnie [run_web_app.bat](X:\CodeX - aplikacja analityczna\run_web_app.bat)

Skrypt:

- utworzy lokalne srodowisko `.venv`
- zainstaluje wymagane biblioteki
- uruchomi aplikacje pod lokalnym adresem `http://localhost:8501`

## Jak otworzyc na innym komputerze w tej samej sieci

Uzyj [run_web_app_lan.bat](X:\CodeX - aplikacja analityczna\run_web_app_lan.bat).

Ta wersja uruchamia aplikacje na:

```text
http://0.0.0.0:8501
```

Nastepnie na drugim komputerze otwierasz:

```text
http://IP_KOMPUTERA_HOSTA:8501
```

Przyklad:

```text
http://192.168.0.25:8501
```

## Wymagania

- Windows
- Python 3.11+ zainstalowany i dostepny jako `py` lub `python`
- polaczenie z internetem przy pierwszym uruchomieniu, aby pobrac zaleznosci

## Reczne uruchomienie

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Wdrożenie na GitHub / Streamlit Cloud

Aby mieć poprawną wersję w GitHub i na Streamlit Cloud, repozytorium musi zawierać plik `app.py` oraz `requirements.txt` w katalogu głównym.

`app.py` jest lekkim wrapperem dla `streamlit_app.py`, dzieki czemu Streamlit Cloud moze uruchamiac aplikacje standardowym entrypointem bez duplikowania logiki.

1. Wypchnij kod do repozytorium GitHub, np. `https://github.com/Pjoterrrr/ShipmentAnalyzer`.
2. Skonfiguruj Streamlit Cloud, wybierając to repo i gałąź `main`.
3. W Streamlit Cloud jako entrypoint użyj:

```text
app.py
```

Po wdrożeniu aplikacja powinna być dostępna pod adresem Streamlit Cloud lub pod własną domeną przypisaną w panelu Streamlit.

> Lokalnie aplikacja działa pod `http://localhost:8501`, a na LAN pod `http://0.0.0.0:8501` przy użyciu `run_web_app_lan.bat`.

## Budowanie paczki EXE

W projekcie jest przygotowany launcher:

- [launcher.py](X:\CodeX - aplikacja analityczna\launcher.py:1)

oraz build pod PyInstaller:

- [ShipmentAnalyzerWeb.spec](X:\CodeX - aplikacja analityczna\ShipmentAnalyzerWeb.spec:1)
- [build_exe_webapp.bat](X:\CodeX - aplikacja analityczna\build_exe_webapp.bat:1)

Budowanie:

```powershell
& "X:\CodeX - aplikacja analityczna\build_exe_webapp.bat"
```

Po zbudowaniu gotowa paczka pojawi sie w:

```text
dist\ShipmentAnalyzerWeb\
```

W srodku znajdziesz plik `ShipmentAnalyzerWeb.exe`, ktory uruchamia aplikacje lokalnie i otwiera ja w przegladarce.

Skrypt buduje EXE na lokalnym dysku tymczasowym Windows i dopiero potem kopiuje wynik do `dist\ShipmentAnalyzerWeb`, co omija problemy z pakowaniem bezposrednio na dysku `X:`.

## Ważne

- logo aplikacji jest zapisane lokalnie w `assets/logo.png`, wiec projekt nie zalezy juz od prywatnej sciezki systemowej
- ikona launchera EXE jest zapisana w `assets/icon.ico`
- konfiguracja Streamlit znajduje sie w `.streamlit/config.toml`
- konfiguracja logowania znajduje sie w `config/users.json`
- jesli chcesz przeniesc aplikacje na inny komputer, kopiuj caly katalog razem z `assets`, `.streamlit` i `config`
