# ✈️ GIOVEN TRAVEL — Excel Toolkit

Strumento web per filtrare le prenotazioni da un file Excel e generare
elenchi puliti (Rooming list / Passenger list) in base a un intervallo di date.

App realizzata con [Streamlit](https://streamlit.io).

## File del progetto

| File | Cosa fa |
|------|---------|
| `streamlit_app.py` | L'applicazione (interfaccia + logica) |
| `requirements.txt` | Le librerie Python necessarie |
| `.streamlit/config.toml` | Il tema colori (viola/rosa) |

## Come metterla online (link usabile da PC e telefono)

Deploy gratuito con **Streamlit Community Cloud**:

1. Vai su **https://share.streamlit.io** e accedi con il tuo account GitHub.
2. Clicca **"Create app"** → **"Deploy a public app from GitHub"**.
3. Compila:
   - **Repository:** `davechen1209/miki`
   - **Branch:** `main` (dopo aver unito le modifiche) oppure il branch di sviluppo
   - **Main file path:** `streamlit_app.py`
4. Clicca **"Deploy"** e attendi qualche minuto.
5. Ottieni un link tipo `https://<nome>.streamlit.app` — aprilo dove vuoi. ✅

## Provarla sul proprio computer

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Formato del file Excel

A seconda della modalità scelta, il file deve contenere queste colonne:

- **排房表 (Rooming list):** `首晚入住日期`, `末晚入住日期`
- **团员名单 (Passenger list):** `上团日期`, `下团日期`

Le date scritte male, vuote o illeggibili vengono ignorate automaticamente.
