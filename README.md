# Raport analiști Yahoo Finance

Instalare:

```powershell
py -m pip install -r requirements.txt
```

Rulare:

```powershell
py yahoo_analyst_report.py AAPL
```

Poți analiza mai multe companii într-o singură rulare și vei primi la final un tabel de comparație:

```powershell
py yahoo_analyst_report.py AAPL MSFT NVDA TSLA
```

Alternativ, creează un fișier `simboluri.txt` cu un simbol pe fiecare rând (liniile pot avea comentarii după `#`) și rulează:

```powershell
py yahoo_analyst_report.py --file simboluri.txt
```

La finalul fiecărei rulări, programul creează automat un fișier Excel în dosarul `Output` din folderul din care rulezi comanda. Numele conține data interogării și simbolurile, de exemplu `13.09.26 NKE.xlsx`. Fișierul include o foaie **Sumar** și câte o foaie pentru fiecare companie.

Înlocuiește `AAPL` cu simbolul companiei dorite (de exemplu `MSFT`, `TSLA`, `NVDA`). Programul cere date publice Yahoo Finance în momentul rulării, afișează consensul de recomandări și încearcă să extragă lista **Top Analysts** din pagina Yahoo. Dacă Yahoo nu livrează acea secțiune în format accesibil (se poate întâmpla din cauza schimbărilor de pagină sau a limitării temporare), raportul continuă cu consensul disponibil și explică acest lucru.

Raportul mai include evoluția pe patru luni a recomandărilor: **Strong Buy**, **Buy**, **Hold**, **Underperform** și **Sell**. Acestea sunt valorile din graficul „Analyst Recommendations” Yahoo; `Underperform` este coloana tehnică Yahoo `sell`, iar `Sell` este `strongSell`.

Pentru fiecare simbol sunt afișate și cele mai recente ratinguri disponibile, preluate dinamic din pagina Yahoo: **Date**, **Analyst**, **Rating Action**, **Rating**, **Price Action** și **Price Target**. Dacă Yahoo nu publică aceste câmpuri în sesiunea curentă, programul folosește istoricul standard disponibil prin Yahoo/yfinance ca rezervă. Raportul conține și o scară **Price Targets**: Low, Current, Average și High. Marcajul `●` este prețul curent, iar `◆` este ținta medie.

Programul folosește `yfinance` pentru consens și țintele de preț, deoarece endpoint-ul public direct Yahoo poate răspunde cu eroarea `401 Unauthorized`.

Datele sunt informative și nu reprezintă consultanță financiară.
