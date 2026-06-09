# Planerad Tidplan - Gantt (Mall)

Nedan finns en mall i Mermaid-format som kan visas i er HTML-preview.
Du kan kopiera blocket och byta namn, datum och varaktighet.

```mermaid
gantt
    title Planerad tidplan - Exempel
    dateFormat  YYYY-MM-DD
    axisFormat  %Y-%m-%d
    excludes    weekends

    section Forberedelse
    Krav och scope            :done, prep1, 2026-06-03, 5d
    Teknisk design            :active, prep2, after prep1, 4d

    section Utveckling
    Implementering steg 1     :dev1, after prep2, 7d
    Implementering steg 2     :dev2, after dev1, 7d

    section Test och kvalitet
    Systemtest                :test1, after dev2, 5d
    Buggrattning              :test2, after test1, 4d

    section Leverans
    UAT                       :uat1, after test2, 3d
    Produktion                :milestone, prod1, after uat1, 1d
```

## Snabbguide

1. Byt aktiviteter i varje section.
2. Satt startdatum (YYYY-MM-DD) eller anvand after aktivitet_id.
3. Andra varaktighet med d (dagar), t.ex. 10d.
4. Anvand milestone for viktiga leveranspunkter.
