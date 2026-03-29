# JSON-First Assessment Engine voor Denjoy

## Doel

De huidige Denjoy assessment-engine is sterk in dataverzameling en HTML-rapportage, maar minder geschikt als directe bron voor het portal. Daarom wordt de assessment-laag omgebouwd naar een **JSON-first model**:

- bestaande `psm1` modules blijven bestaan
- iedere module levert naast globale fase-data ook portalvriendelijke payloads
- één orchestrator schrijft JSON-bestanden weg per hoofdstuk en subhoofdstuk
- backend en portal kunnen die JSON-bestanden direct als snapshotbron gebruiken

## Waarom dit beter is dan één groot monolithisch script

Niet aanbevolen:

- alles samenvoegen tot één grote `Invoke-QUBEFoundationScan.ps1`-achtige scan

Wel aanbevolen:

- modulaire `psm1`-modules behouden
- uniforme JSON-outputcontracten toevoegen
- orchestratie en export erboven leggen

Voordelen:

- beter testbaar
- beter herbruikbaar voor live en assessment
- betere aansluiting op hoofdstuk/subhoofdstuk in portal
- minder kans op regressies in bestaande rapportage

## Gewenst patroon per module

Iedere fase-module krijgt uiteindelijk drie verantwoordelijkheden:

1. **assessment data ophalen**
2. **globale PhaseXData vullen**
3. **portal payloads teruggeven of exporteren**

Voorbeeld:

- `Invoke-Phase1Assessment`
- `Get-Phase1PortalPayloads`
- `Export-Phase1PortalJson`

## JSON outputcontract

Per hoofdstuk/subhoofdstuk:

```json
{
  "section": "gebruikers",
  "subsection": "users",
  "label": "Gebruikers",
  "source": "assessment",
  "generated_at": "2026-03-27T12:00:00Z",
  "assessmentId": "20260327-120000",
  "summary": {},
  "items": [],
  "meta": {
    "notes": [],
    "permissions": []
  }
}
```

## Voorbeeldmapping voor Phase 1

`Phase1-UsersLicensing.psm1` levert als referentie:

- `gebruikers.users`
- `gebruikers.licenses`
- `identity.mfa`

Dit sluit direct aan op bestaande portalwerkruimtes.

## Exportorchestratie

De centrale exportmodule:

- [Export-AssessmentJson.psm1](/Users/demac/Downloads/Denjoy-it-site-main-2/assessment/Modules/Export-AssessmentJson.psm1)

doet nu:

- outputmap maken
- fasegebonden JSON-export starten
- manifestbestand schrijven

## Integratie in de assessment-runner

De bestaande assessment-runner krijgt een extra optie:

- `-ExportJson`

Daardoor kan één assessment-run nu opleveren:

- HTML
- CSV
- portal JSON

## Aanbevolen vervolgstappen

1. Phase 2 volgens hetzelfde patroon ombouwen
2. manifest verrijken met staleness- en source-metadata
3. backend JSON-import laten gebruiken als officiële assessment snapshotbron
4. daarna pas generiek uitbreiden naar Phase 3, 4, 5 en 6

## Conclusie

Het idee is sterk, zolang het **modulair** wordt uitgevoerd. De juiste aanpak voor Denjoy is:

- niet één grote scanfile centraal zetten
- wel elke bestaande module uitrusten met een uniform JSON-contract
- en dat via een centrale exportlaag laten uitspugen voor het portal
