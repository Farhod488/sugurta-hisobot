# Sug'urta hisobot — Тушум

Desktop-приложение для формирования трёх отчётов **Тушум** (DOCX) по данным KUZATUV и OSAGO.
Также есть Streamlit-модуль для сводного Excel-отчёта.

## Входные файлы (Тушум)

Все файлы **необязательны по отдельности**. Нужен хотя бы один.

| Файл | Что обновляет | Если файла нет |
|------|---------------|----------------|
| `KUZATUV *.pdf` | абзац 4 (итоги с начала) | абзац не трогаем |
| `KUZATUV.xlsx` | абзац 5 (период) | абзац не трогаем |
| `OSAGO *.pdf` | абзац 8 | абзац не трогаем |
| `OSAGO.xlsx` | абзац 9 | абзац не трогаем |

Итоговый абзац (MBL5+MBL9) обновляется только если загружены **оба** Excel.

## Три отчёта

1. **Ш.Арипов** — 4 компании KUZATUV + 2 компании OSAGO  
2. **Ш.Абдимуродов** — 3 компании KUZATUV + 1 компания OSAGO  
3. **Б.Хамракулов** — свод по всем компаниям за период  

Период (например `24-31-avgust`) задаётся в интерфейсе и подставляется в абзацы 5, 9 и итоговый.

## Установка и запуск

```powershell
cd D:\Projects\sugurta-hisobot
python -m venv .venv
.\.venv\Scripts\pip install -r requirements-desktop.txt
.\run.bat
```

или:

```powershell
.\.venv\Scripts\python.exe main.py
```

### Streamlit (Excel-сводка)

Локально:

```powershell
.\.venv\Scripts\pip install -r requirements.txt
.\.venv\Scripts\streamlit.exe run app.py
```

На Streamlit Cloud главный файл: `app.py` (зависимости из `requirements.txt`).

## Логика (Арипов)

- Абзацы 1–3, 6–7 — без изменений  
- Абзац 4 — `Jami` из KUZATUV.pdf (полисы → «1 mln 345 ming…», суммы mln→mlrd)  
- Абзац 5 — период + суммы по 4 компаниям из KUZATUV.xlsx  
- Абзац 8 — `Jami` из OSAGO.pdf  
- Абзац 9 — период + суммы по 2 компаниям из OSAGO.xlsx  
- Абзац 10 — MBL(абз.5) + MBL(абз.9)

Шаблоны: `templates/tushum_*.docx`  
Примеры входов: `templates/source/`  
Готовые файлы: `output/`
