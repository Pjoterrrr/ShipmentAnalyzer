# Release Dashboard Online

Profesjonalna aplikacja Streamlit do porównywania dwóch plików Excel:

- poprzedni release / poprzedni plan
- aktualny release / aktualny plan

Aplikacja zachowuje logikę analityczną biznesową:

- porównanie `previous` vs `current`
- KPI i alerty przy `abs(Percent Change) >= 15`
- wykres trendu, delta, struktura zmian i raport produktu
- macierz release'u
- eksport CSV i Excel
- logowanie użytkowników

Projekt jest przygotowany do:

- pracy lokalnej
- hostowania na GitHub
- wdrożenia na Streamlit Community Cloud
- aktualizacji przez zwykły `git push`

## Struktura projektu

```text
.
|-- .streamlit/
|   `-- config.toml
|-- app.py
|-- release_dashboard_updated.py
|-- streamlit_app.py
|-- requirements.txt
|-- README.md
|-- .gitignore
|-- secrets_example.toml
|-- assets/
|   `-- .gitkeep
`-- config/
    `-- users.example.json
```

## Plik startowy

Do uruchamiania lokalnego i deployu na Streamlit Community Cloud rekomendowany jest:

- `app.py`

Dodatkowo działają także:

- `release_dashboard_updated.py`
- `streamlit_app.py`

## Logowanie

Aplikacja wspiera dwa źródła logowania:

1. `config/users.json` - preferowane lokalnie
2. `st.secrets` - preferowane na Streamlit Community Cloud

Priorytet jest następujący:

1. jeśli istnieje `config/users.json`, aplikacja używa pliku lokalnego
2. jeśli nie ma pliku lokalnego, aplikacja sprawdza `st.secrets`
3. jeśli nie ma żadnej konfiguracji, ekran logowania pokaże prosty komunikat konfiguracyjny

### Lokalny plik użytkowników

Skopiuj:

- `config/users.example.json`

do:

- `config/users.json`

Przykład:

```json
{
  "users": [
    {
      "username": "planner",
      "display_name": "Planner",
      "role": "Analyst",
      "active": true,
      "salt": "REPLACE_WITH_HEX_SALT",
      "password_hash": "REPLACE_WITH_HEX_PASSWORD_HASH"
    }
  ]
}
```

### Sekrety dla Streamlit Community Cloud

Użyj pliku `secrets_example.toml` jako wzoru.

Przykład:

```toml
[[auth.users]]
username = "planner"
display_name = "Planner"
role = "Analyst"
active = true
salt = "REPLACE_WITH_HEX_SALT"
password_hash = "REPLACE_WITH_HEX_PASSWORD_HASH"
```

### Jak wygenerować `salt` i `password_hash`

```bash
python -c "import os,binascii,hashlib,getpass; pwd=getpass.getpass('Password: '); salt=os.urandom(16); h=hashlib.pbkdf2_hmac('sha256', pwd.encode('utf-8'), salt, 120000); print('salt=' + binascii.hexlify(salt).decode()); print('password_hash=' + binascii.hexlify(h).decode())"
```

Wygenerowane wartości wklej do:

- `config/users.json`
albo
- `.streamlit/secrets.toml`

## Logo

Logo jest opcjonalne.

Jeśli chcesz je wyświetlać:

1. dodaj plik `assets/logo.png`
2. uruchom aplikację ponownie

Jeśli pliku nie ma, dashboard działa normalnie bez logo.

## Uruchomienie lokalne

1. Utwórz i aktywuj środowisko wirtualne.
2. Zainstaluj zależności:

```bash
pip install -r requirements.txt
```

3. Skonfiguruj logowanie:

- opcja A: utwórz `config/users.json` na podstawie `config/users.example.json`
- opcja B: przygotuj `.streamlit/secrets.toml` na podstawie `secrets_example.toml`

4. Opcjonalnie dodaj `assets/logo.png`

5. Uruchom aplikację:

```bash
streamlit run app.py
```

Alternatywnie:

```bash
streamlit run release_dashboard_updated.py
```

## GitHub

Jeśli tworzysz nowe repozytorium:

```bash
git init
git add .
git commit -m "Prepare Streamlit dashboard"
git branch -M main
git remote add origin https://github.com/TWOJ_LOGIN/TWOJE_REPO.git
git push -u origin main
```

Nie commituj:

- `.streamlit/secrets.toml`
- `config/users.json`

Oba pliki są już ujęte w `.gitignore`.

## Deploy na Streamlit Community Cloud

1. Wejdź na `https://share.streamlit.io`
2. Zaloguj się przez GitHub
3. Kliknij `Create app`
4. Wybierz repozytorium
5. Ustaw branch `main`
6. Ustaw plik startowy `app.py`
7. Otwórz `Advanced settings`
8. Wklej sekrety w formacie z `secrets_example.toml`
9. Kliknij `Deploy`

Po wdrożeniu aplikacja będzie dostępna pod publicznym linkiem `https://...streamlit.app`.

## Aktualizowanie aplikacji

Po każdej zmianie:

```bash
git add .
git commit -m "Update dashboard"
git push
```

Streamlit Community Cloud pobierze zmiany z GitHub i odświeży aplikację.

## Wskazówki wdrożeniowe

- przechowuj prawdziwe dane logowania wyłącznie w `st.secrets` lub lokalnym pliku ignorowanym przez Git
- uruchamiaj aplikację z katalogu głównego repozytorium
- jeśli dodasz nowe biblioteki, zaktualizuj `requirements.txt`

## Szybka checklista

1. `pip install -r requirements.txt`
2. skonfiguruj użytkowników w `config/users.json` lub `.streamlit/secrets.toml`
3. `streamlit run app.py`
4. wrzuć repo na GitHub
5. podłącz repo do Streamlit Community Cloud
6. dodaj sekrety w panelu Cloud
7. kliknij `Deploy`
