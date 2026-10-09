# Tilbudsavis-automation

Programmet finder tilbud fra den valgte ugeavis pr. butik. Det bruger offentlige
eTilbudsavis-sider for de øvrige kæder og Coops officielle avisvisninger for
365discount, Brugsen, Kvickly og SuperBrugsen. Tilbuddene matches til avisens
publikations-ID, og avisens faktiske forside vises i PowerPointen. Reglerne står
i `config/categories.json`. Søgte tilbud, der ikke matcher en kategori, gemmes
separat i `mulige-tilbud-til-kategorisering.csv`.

## Kør lokalt

Kræver Python 3.10 eller nyere.

```bash
python -m pip install -r requirements.txt
python src/offer_automation.py
```

Uden `--week` bruges næste ISO-uge. For at lave præsentationen specifikt for uge
42 køres:

```bash
python src/offer_automation.py --week 42
```

Filerne bliver lagt i `output/`. Søgeord, brands, butiknavne og sidernes slugs
kan redigeres i `config/categories.json`.

## Kør via GitHub

Workflowet under Actions kan startes manuelt med **Run workflow**. Det kører
også automatisk hver torsdag kl. 07:00 UTC. Når kørslen er færdig, hentes
PowerPoint-filen og CSV-filerne fra workflowets **Artifacts**. GitHub gemmer dem
i 30 dage.

## Begrænsninger i første version

- Søgningen bruger offentlige eTilbudsavis-sider og kategoriernes søgeord. Siden
	viser højst 30 avisresultater pr. søgning, og ukendte produktnavne bliver ikke
	automatisk kategoriseret. Brug `mulige-tilbud-til-kategorisering.csv` til at
	finde forslag til nye søgeord og brands.
- Hvis eTilbudsavis endnu ikke har publiceret en avis for måluge for en butik,
	udelades dens tilbud, og forsiden viser, at ugeavisen ikke kunne findes.
- Coop-aviser vælges ud fra deres gyldighedsdatoer. Et katalog med en ugeetiket
	som spænder over flere uger kan derfor være valgt, når det gælder i måluge.
- Når en butik har flere aviser mærket med samme uge, vælges den senest
	publicerede. Det holder udvælgelsen til én avis pr. butik.
- Tilbudsreglerne er bevidst simple. Gennemse CSV-filen og tilføj manglende
	mærker eller søgeord i konfigurationen.
- eTilbudsavis kan ændre sidernes struktur eller begrænse automatiske
	forespørgsler. Workflowet bruger en pause mellem søgninger, men den faste
	automatisering bør kun bruges i overensstemmelse med Tjeks vilkår.
- Eksempelpræsentationen var ikke tilgængelig i projektet, så layoutet er et
	første udkast og ikke en præcis kopi.
