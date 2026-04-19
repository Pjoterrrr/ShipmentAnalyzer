# Release Dashboard Online

Profesjonalny dashboard Streamlit do porównywania dwóch plików Excel:

- previous release / previous plan
- current release / current plan

Aplikacja zachowuje całą logikę analityczną:

- porównanie `previous` vs `current`
- KPI
- alerty przy `abs(Percent Change) >= 15`
- trend, delta, change mix, raport produktu
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
├── .streamlit/
│   └── config.toml
├── app.py
├── release_dashboard_updated.py
├── streamlit_app.py
├── requirements.txt
├── README.md
├── .gitignore
├── secrets_example.toml
├── assets/
│   └── .gitkeep
└── config/
    └── users.example.json
```

## Entry point

Do deployu online rekomendowany jest:

- `app.py`

Działają też:

- `release_dashboard_updated.py`
- `streamlit_app.py`

## Logowanie

Aplikacja obsługuje dwa źródła logowania:

1. `st.secrets` - preferowane dla Streamlit Community Cloud
2. `config/users.json` - fallback lokalny

Jeżeli żadne źródło nie jest skonfigurowane, aplikacja nie crashuje, tylko pokazuje komunikat konfiguracyjny na ekranie logowania.

### Format użytkowników w `st.secrets`

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

### Format użytkowników lokalnych

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

### Jak wygenerować salt i hash hasła

Uruchom lokalnie:

```bash
python -c "import os,binascii,hashlib,getpass; pwd=getpass.getpass('Password: '); salt=os.urandom(16); h=hashlib.pbkdf2_hmac('sha256', pwd.encode('utf-8'), salt, 120000); print('salt=' + binascii.hexlify(salt).decode()); print('password_hash=' + binascii.hexlify(h).decode())"
```

Wygenerowane wartości wklej do:

- `.streamlit/secrets.toml`
albo
- `config/users.json`

## Logo

Logo jest opcjonalne.

Jeśli chcesz je pokazać:

1. dodaj plik:
   - `assets/logo.png`
2. uruchom aplikację ponownie

Jeżeli pliku nie ma, dashboard działa normalnie bez logo.

## Uruchomienie lokalne

1. Utwórz i aktywuj środowisko wirtualne.
2. Zainstaluj zależności:

```bash
pip install -r requirements.txt
```

3. Skonfiguruj logowanie:

- opcja A: skopiuj `secrets_example.toml` do `.streamlit/secrets.toml`
- opcja B: skopiuj `config/users.example.json` do `config/users.json`

4. Opcjonalnie dodaj `assets/logo.png`

5. Uruchom aplikację:

```bash
streamlit run app.py
```

Alternatywnie:

```bash
streamlit run release_dashboard_updated.py
```

## Przygotowanie repozytorium GitHub

1. Załóż nowe repozytorium na GitHub.
2. Skopiuj do niego wszystkie pliki projektu.
3. Upewnij się, że nie commitujesz:
   - `.streamlit/secrets.toml`
   - `config/users.json`
4. Zainicjalizuj repo lokalnie:

```bash
git init
git add .
git commit -m "Prepare Streamlit dashboard for Community Cloud deploy"
git branch -M main
git remote add origin https://github.com/TWOJ_LOGIN/TWOJE_REPO.git
git push -u origin main
```

## Deploy na Streamlit Community Cloud

1. Wejdź na:
   - `https://share.streamlit.io`
2. Zaloguj się przez GitHub.
3. Kliknij `Create app`.
4. Wybierz repozytorium GitHub.
5. Ustaw branch:
   - `main`
6. Ustaw plik startowy:
   - `app.py`
7. Otwórz `Advanced settings`.
8. Wklej zawartość `secrets_example.toml` po podmianie `salt` i `password_hash` na prawdziwe wartości.
9. Kliknij `Deploy`.

Po wdrożeniu aplikacja będzie dostępna pod publicznym linkiem `https://...streamlit.app`.

## Aktualizowanie aplikacji

Po każdej zmianie w kodzie:

```bash
git add .
git commit -m "Update dashboard"
git push
```

Community Cloud pobierze zmiany z GitHub i odświeży aplikację online.

## Wskazówki deploymentowe

- trzymaj prawdziwe dane logowania wyłącznie w `st.secrets` lub lokalnym pliku ignorowanym przez Git
- uruchamiaj lokalnie komendą z katalogu głównego repo
- nie commituj żadnych plików z prawdziwymi hasłami
- jeśli dodasz nowe biblioteki, zaktualizuj `requirements.txt`

## Szybka checklista

1. `pip install -r requirements.txt`
2. skonfiguruj użytkowników w `.streamlit/secrets.toml` lub `config/users.json`
3. `streamlit run app.py`
4. wrzuć repo na GitHub
5. podłącz repo do Streamlit Community Cloud
6. dodaj sekrety w panelu Cloud
7. kliknij `Deploy`
