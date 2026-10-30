# Klickern

Telefonen och hörlurarna som clicker till PowerPoint.

Du går där du vill i rummet. Telefonen visar dina anteckningar och vilken bild du är på, och du byter bild med tummen eller med hörlurarnas knappar. Ingen installation och inget konto. Gratis, för Windows.

**[Hämta Klickern.exe](https://github.com/Pluggentipsar/klickern/releases/latest/download/Klickern.exe)** (cirka 86 MB)

## Kom igång

1. Hämta `Klickern.exe` och lägg den där du hittar den, till exempel på skrivbordet.
2. Starta filen. Windows varnar första gången, eftersom programmet inte är kodsignerat: klicka på **Mer information** och sedan på **Kör ändå**.
3. Ett fönster och en sida i webbläsaren öppnas. Skanna QR-koden med telefonens kamera och öppna länken. Låt fönstret vara öppet medan du föreläser.
4. Tryck på **Styr med hörlurarna** i telefonen om du vill använda hörlurarnas knappar.
5. Starta bildspelet i PowerPoint som vanligt, eller från telefonen.

Stäng fönstret för att stänga av Klickern.

## Det här får du

- En stor knapp för Nästa under tummen, och en liten för Bakåt som går att gömma.
- Dina anteckningar från PowerPoint, i tre textstorlekar.
- Prickar som visar animationsklicken på bilden.
- Hopp till ett avsnitt eller en bild, utan att publiken ser någon meny.
- Svart skärm.
- Hörlurarna: spela/pausa och nästa spår går framåt, förra spåret bakåt. Ett dubbeltryck går snabbast, eftersom hörlurarna väntar en stund efter ett enkeltryck.

PowerPoint behöver inte ligga överst. Utan bildspel skickar knapparna piltangenter till fönstret som ligger överst.

## Så når telefonen datorn

**Samma wifi** är enklast: inget konto och inget mer att installera. Många skol- och konferensnät stänger dock av trafik mellan enheter, och då laddar sidan aldrig i telefonen. Trafiken är okrypterad, så använd det på nät du litar på. Första gången kan Windows fråga om brandväggen: välj **Tillåt**.

**[Tailscale](https://tailscale.com/download)** är en privat, krypterad förbindelse mellan dina egna enheter, och fungerar på vilket nät som helst. Installera det på datorn och i telefonen och logga in med samma konto. Välj sedan fliken **Tailscale** på Klickerns sida och tryck på **Slå på Tailscale för Klickern**. Första gången ber Tailscale dig att slå på https: öppna länken, godkänn och tryck på knappen igen.

## Skydd

- Telefonens adress innehåller en slumpad nyckel på 128 bitar. **Ny länk** på datorsidan byter den.
- Datorsidan, med QR-koden, lyssnar bara på den egna datorn.
- Ett steg kräver sidans eget ursprung och en egen rubrik, så en annan webbplats i samma webbläsare kan inte styra.
- Dina anteckningar går bara mellan din dator och din telefon. Klickern har ingen tjänst på nätet.

## Begränsningar

- Bara Windows. Klickern styr PowerPoint genom dess automationsgränssnitt (COM), som bara finns där.
- Provat med Android och hörlurarna Nero TX2. iPhone är inte provad.
- Programmet är inte kodsignerat.

## Bygg själv

Kräver Node 24 på Windows.

```bash
npm install
npm start        # kör programmet direkt ur mappen
npm run bygg     # bygger dist/Klickern.exe
```

| Fil | Vad den gör |
|---|---|
| `program.mjs` | Programmet: nyckeln, wifi, Tailscale, datorsidan och start av webbläsaren. |
| `klickern.mjs` | Servern: telefonsidan, läget som ström och stegen. |
| `powerpoint.ps1` | Bryggan till PowerPoint. Måste sparas med BOM. |
| `telefon.html` | Telefonsidan. |
| `dator.html` | Datorsidan med QR-koden. |
| `bygg.mjs` | Bygger exe-filen med Nodes single executable application. |

Prova inte med en PowerPoint som någon annan har öppen: kommandona `start` och `ga N` startar ett bildspel. `KLICKERN_TORRT=1` hindrar piltangenterna från att skickas.

## Licens

[MIT](LICENSE), av Joel Rangsjö. Exe-filen innehåller Node.js och npm-paket med egna licenser, se [THIRD-PARTY-NOTICES](THIRD-PARTY-NOTICES.md).

Klickern är gjord av [Joel Rangsjö](https://joelsai.com), som också har gjort [Presenter](https://github.com/Pluggentipsar/presenter), ett program för föreläsningar som rör sig.
