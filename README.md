# Podpis Allrisk – Outlook doplněk (v2)

Vkládá firemní podpis Allrisk při psaní nové zprávy, odpovědi i přeposlání.
Údaje bere přímo z Microsoft 365 (Graph `/me`) pod přihlášením uživatele – žádný export, žádná jména v repozitáři.

## Jak to funguje

1. Outlook spustí `onNewMessageCompose` (nebo `onMessageFromChanged` při přepnutí pole Od).
2. Když je adresa v poli Od v doméně z `DOMAINS` (`src/config.js`), doplněk:
   - vezme profil z mezipaměti (max. 4 h), jinak si potichu (NAA) vyžádá token a načte `/me` z Graphu,
   - sestaví podpis podle `src/template.js` a vloží ho (`setSignatureAsync`),
   - nastaví hlavičku `X-Allrisk-Signature: addin`, podle které transport rule podpis nepřidá podruhé.
3. Když tiché přihlášení selže, nevloží nic (podpis dodá transport rule) a v pilotu ukáže lištu „Otevřít“ → panel doplňku → **Přihlásit a vložit podpis**.

| Údaj v podpisu | Pole v Entra ID |
|---|---|
| Jméno | `displayName` (jinak `givenName` + `surname`) |
| Funkce | `jobTitle` |
| Telefon | `mobilePhone`, jinak `businessPhones[0]` |
| E-mail | adresa z pole Od; ze sdílené schránky osobní e-mail pisatele |
| Adresa | `streetAddress`, `postalCode`, `city` (prázdné → sídlo Brno) |
| Fotka | `onPremisesExtensionAttributes.extensionAttribute15` |

`disableClientSignatureAsync` se záměrně nepoužívá – v klasickém Outlooku a na Macu trvale vypne výchozí podpis účtu.

## Úpravy

- `src/config.js` – Client ID, Tenant ID, domény, hlavička
- `src/template.js` – vzhled podpisu (logo, banner, sociální sítě)
- `manifest.xml` – mění se jen při změně adresy nebo událostí (pak zvýšit `<Version>` a znovu nahrát do Integrated apps)

Po `git push` do `main` GitHub Actions sám spustí `npm ci && npm run build` a nasadí `dist/` na
`https://nikitasei.github.io/allrisk-signature/`.

Lokální sestavení (není nutné): `npm install && npm run build`.

## Závislost: repozitář `nikitasei.github.io`

Klasický Outlook pro Windows smí volat Graph z doplňku jen tehdy, když na **kořeni domény** leží
`https://nikitasei.github.io/.well-known/microsoft-officeaddins-allowed.json` se seznamem povolených skriptů.
Ten soubor je v samostatném repozitáři `nikitasei.github.io` (uživatelská stránka GitHub Pages).
