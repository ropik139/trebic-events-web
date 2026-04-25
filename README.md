# Trebic events web

Jednoduchy web pro sledovani kulturnich akci v Trebici a okoli do 10 km.

## Co projekt dela

- denne nacita kulturni akce z nastavenych zdroju
- jako sekundarni zdroj umi vyuzit i tydenni prehledy z webu mesta Trebic
- filtruje jen akce v okruhu 10 km od Trebice
- zobrazuje horizont 14 dni dopredu
- odstranuje uz uplynule akce
- generuje mobilne pouzitelny web do `docs/index.html`
- umi publikaci na GitHub Pages

## Hlavni soubory

- `scripts/update-trebic-events.ps1` - stahne a zpracuje akce
- `scripts/publish-trebic-events-site.ps1` - pripravi web do slozky `docs`
- `.github/workflows/trebic-events-pages.yml` - denni GitHub Actions workflow
- `trebic-events.settings.json` - konfigurace zdroju a vystupu

## Lokalni spusteni

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\update-trebic-events.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\publish-trebic-events-site.ps1
```

## GitHub Pages

Na GitHubu nastav v `Settings -> Pages` zdroj `GitHub Actions`.
Workflow pak web sam prepocita a nasadi.
