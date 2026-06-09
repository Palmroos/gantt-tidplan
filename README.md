# Gantt Planning Repo

Detta repo ar till for planerade tidplaner och Gantt-scheman.

## Innehall

- planned_timeline_gantt_sv.md: Mallar och exempel i Mermaid Gantt.

## Anvandning

1. Redigera planfilen.
2. Generera Gantt-filer fran Excel:

   ```powershell
   .\.venv\Scripts\python.exe .\build_gantt.py
   ```

3. Visa Mermaid-diagram i VS Code markdown preview eller HTML-preview.
4. Versionera andringar med git.
