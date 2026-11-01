# MS Project → SharePoint Online sync add-in

VSTO doplnok pre **Microsoft Project desktop** (Project Professional / Project pre Microsoft 365), ktorý pri každom
**uložení projektu** prenesie dáta otvoreného projektu do **SharePoint Online listov**.

- projekt, úlohy, zdroje a priradenia → 4 SharePoint listy (názvy si nastavíte)
- listy a stĺpce sa pri prvej synchronizácii **vytvoria automaticky**
- nahrávajú sa iba **zmenené** položky (hash obsahu), odstránené úlohy sa zo SharePointu zmažú
- uloženie v MS Project sa **nezdrží** — dáta sa načítajú pri uložení a odošlú na pozadí
- prihlásenie cez Microsoft Entra ID (MSAL), token sa ukladá šifrovane (DPAPI), prihlasujete sa iba raz
- ribbon karta **SharePoint**: *Sync now*, *Sync on save*, *Open site*, *Settings*, *Log*

## Architektúra

```
ProjectSharePointSync.sln
└─ src/
   ├─ ProjectSharePointSync.Core/        knižnica (net48, bez závislosti na Office)
   │  ├─ Auth/GraphTokenProvider.cs      MSAL – delegované tokeny pre Microsoft Graph
   │  ├─ SharePoint/GraphClient.cs       REST + JSON batching (20 req/batch) + retry pri 429/503
   │  ├─ SharePoint/ListProvisioner.cs   nájde/vytvorí listy a stĺpce, indexuje kľúčové stĺpce
   │  ├─ SharePoint/ProjectSyncService.cs  diff (pridať / upraviť / zmazať) a zápis
   │  ├─ Schema/                         definícia listov a stĺpcov
   │  └─ Configuration/                  nastavenia (%APPDATA%\ProjectSharePointSync\settings.json)
   └─ ProjectSharePointSync.AddIn/       VSTO doplnok pre MS Project
      ├─ ThisAddIn.cs                    udalosť Application.ProjectBeforeSave
      ├─ SyncCoordinator.cs              fronta synchronizácií na pozadí, notifikácie
      ├─ Reading/ProjectReader.cs        čítanie úloh/zdrojov/priradení cez COM object model
      ├─ Ribbon/                         ribbon XML + callbacky
      └─ UI/SettingsForm.cs              okno nastavení
```

Priebeh pri uložení:

1. MS Project vyvolá `ProjectBeforeSave`.
2. Doplnok na UI vlákne prečíta projekt do pamäťovej kópie (COM objekty sa nesmú používať z iného vlákna).
3. Na pozadí: nájde/vytvorí listy, načíta existujúce položky projektu (`PsProjectId`), porovná hash
   a cez Graph `$batch` pošle iba rozdiely.
4. Výsledok sa zobrazí v stavovom riadku MS Projectu, chyby v dialógu a v logu
   (`%LOCALAPPDATA%\ProjectSharePointSync\logs`).

Ak MS Project zatvoríte počas prebiehajúcej synchronizácie, doplnok počká až 20 s, kým sa dokončí.

### Prečo Microsoft Graph a nie CSOM

Balík `Microsoft.SharePointOnline.CSOM` pre .NET Framework pribaľuje vlastnú (starú) verziu
`Microsoft.Identity.Client.dll`, ktorá koliduje s aktuálnym MSAL. Graph REST API nemá žiadnu ďalšiu
závislosť a JSON batching znižuje počet požiadaviek 20×.

## Identifikácia projektu

Pri prvej synchronizácii sa do .mpp súboru uloží vlastnosť dokumentu **`PsSyncProjectId`** (GUID).
Podľa nej sa projekt páruje so záznamami v SharePointe, takže premenovanie alebo presun súboru nevadí.

> Pozor: ak súbor skopírujete (Uložiť ako…) a chcete z kópie samostatný projekt, zmažte v kópii
> vlastnosť `PsSyncProjectId` (Súbor → Informácie → Rozšírené vlastnosti → Vlastné).

Voliteľná vlastnosť **`PsSyncSiteUrl`** (nastaviteľná v okne *Settings* → *Site URL for this project*)
presmeruje konkrétny projekt na inú SharePoint lokalitu.

## SharePoint listy

Všetky vlastné stĺpce majú prefix `Ps` (internal name). Spoločné stĺpce vo všetkých listoch:

| Stĺpec | Typ | Popis |
|---|---|---|
| `Title` | text | názov projektu / úlohy / zdroja |
| `PsProjectId` | text, indexovaný | ID projektu |
| `PsKey` | text, indexovaný | jedinečný kľúč záznamu (`{ProjectId}:T:{UniqueID}` …) |
| `PsHash` | text | hash obsahu – na detekciu zmien |

**MSP Projects** – `PsFilePath, PsStart, PsFinish, PsDurationDays, PsWorkHours, PsActualWorkHours, PsCost,
PsPercentComplete, PsManager, PsCompany, PsAuthor, PsStatusDate, PsTaskCount, PsResourceCount, PsLastSync, PsLastSyncBy`

**MSP Tasks** – `PsProjectName, PsUniqueId, PsTaskId, PsWbs, PsOutlineLevel, PsParentUniqueId, PsStart, PsFinish,
PsActualStart, PsActualFinish, PsBaselineStart, PsBaselineFinish, PsDeadline, PsDurationDays, PsWorkHours,
PsActualWorkHours, PsCost, PsPercentComplete, PsPercentWorkComplete, PsSummary, PsMilestone, PsCritical, PsActive,
PsResourceNames, PsPredecessors, PsNotes` + voliteľné vlastné polia `PsCf<Názov>` (napr. `PsCfText1`)

**MSP Resources** – `PsProjectName, PsUniqueId, PsResourceId, PsResourceType, PsInitials, PsGroup, PsEmail,
PsMaxUnits, PsStandardRate, PsWorkHours, PsActualWorkHours, PsCost`

**MSP Assignments** – `PsProjectName, PsUniqueId, PsTaskUniqueId, PsResourceUniqueId, PsTaskName, PsResourceName,
PsStart, PsFinish, PsWorkHours, PsActualWorkHours, PsUnits, PsCost, PsPercentWorkComplete`

Jednotky: trvanie v pracovných dňoch (podľa *hodín za deň* v projekte), práca v hodinách,
percentá 0–100, dátumy v UTC.

Viac projektov môže zapisovať do tých istých listov – záznamy sa rozlišujú podľa `PsProjectId`.
Stĺpce `PsProjectId` a `PsKey` sú indexované, takže filtrovanie funguje aj nad 5 000 položiek.

## 1. Registrácia aplikácie v Microsoft Entra ID

1. [Entra admin center](https://entra.microsoft.com) → **App registrations** → **New registration**
   - Name: `MS Project SharePoint Sync`
   - Supported account types: *Accounts in this organizational directory only*
2. **Authentication** → **Add a platform** → **Mobile and desktop applications** →
   vlastné redirect URI **`http://localhost`**.
   Zapnite **Allow public client flows = Yes**.
3. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**:
   - `Sites.Manage.All` – potrebné na automatické vytvorenie listov a stĺpcov
     (ak listy vytvoríte vopred a vypnete *Create missing lists*, stačí `Sites.ReadWrite.All`)
   - `User.Read` (býva pridané automaticky)
4. **Grant admin consent** (odporúčané, inak sa každý používateľ bude pri prvom prihlásení pýtať na súhlas
   alebo ho tenant nepustí).
5. Skopírujte **Application (client) ID** a **Directory (tenant) ID**.

Delegované oprávnenia znamenajú, že doplnok zapisuje vždy **v mene prihláseného používateľa** –
používateľ musí mať na cieľovej lokalite právo upravovať (pri vytváraní listov právo spravovať) listy.

## 2. Build

Požiadavky: Windows, **Visual Studio 2022** s workloadom **Office/SharePoint development**
(obsahuje VSTO a Office PIA), .NET Framework 4.8, nainštalovaný MS Project (2016/2019/2021/Microsoft 365).

1. Otvorte `ProjectSharePointSync.sln`.
2. Pravým na projekt **ProjectSharePointSync.AddIn** → **Properties** → **Signing** →
   **Sign the ClickOnce manifests** → **Create Test Certificate…** (bez certifikátu build zlyhá;
   `.pfx` sa do gitu necommituje).
3. **Build** → NuGet balíky sa obnovia automaticky.
4. **F5** spustí MS Project s načítaným doplnkom (Visual Studio ho zaregistruje pre aktuálneho používateľa).

> Ak by sa VSTO projekt v Visual Studiu nenačítal (napr. iná verzia VSTO nástrojov), vytvorte nový
> projekt **Project VSTO Add-in** (C#, .NET Framework 4.8) s názvom `ProjectSharePointSync.AddIn`,
> nahraďte `ThisAddIn.cs` týmto súborom, pridajte priečinky `Reading`, `Ribbon`, `UI`, súbor
> `SyncCoordinator.cs`, `Ribbon\SyncRibbon.xml` ako *Embedded Resource*, referenciu na projekt
> `ProjectSharePointSync.Core` a NuGet balíky `Microsoft.Identity.Client`,
> `Microsoft.Identity.Client.Extensions.Msal`, `Newtonsoft.Json`.

## 3. Nasadenie používateľom

- **ClickOnce**: projekt → **Publish** → zložka/sieťový share/web. Používateľ spustí `setup.exe`
  (nainštaluje aj VSTO Runtime). Aktualizácie sa kontrolujú každých 7 dní.
- **Firemné nasadenie**: podpíšte manifesty skutočným code-signing certifikátom a nasaďte cez Intune/GPO
  (MSI/`VSTOInstaller.exe /i <url>.vsto /s`).

## 4. Nastavenie

Karta **SharePoint** → **Settings**:

| Nastavenie | Popis |
|---|---|
| SharePoint site URL | napr. `https://contoso.sharepoint.com/sites/pmo` |
| Tenant ID | GUID alebo doména tenanta |
| Application (client) ID | z registrácie aplikácie |
| Site URL for this project | voliteľne iná lokalita pre aktívny projekt |
| Synchronize automatically when saved | synchronizácia pri uložení (aj tlačidlo *Sync on save*) |
| Tasks / Resources / Assignments | čo sa má prenášať |
| Create missing lists and columns | automatické vytvorenie listov a stĺpcov |
| Delete list items removed from the project | mazanie položiek, ktoré v projekte už nie sú |
| Lists | názvy listov |
| Extra task fields | ďalšie polia úloh (napr. `Text1`, `Number2`, `Flag1`, názov vlastného poľa), uložia sa ako text |

**Test connection** overí prihlásenie a prístup k lokalite. Nastavenia sa ukladajú do
`%APPDATA%\ProjectSharePointSync\settings.json` – súbor môžete pripraviť centrálne a rozdistribuovať.

Príklad `settings.json`:

```json
{
  "SiteUrl": "https://contoso.sharepoint.com/sites/pmo",
  "TenantId": "contoso.onmicrosoft.com",
  "ClientId": "00000000-0000-0000-0000-000000000000",
  "SyncOnSave": true,
  "SyncTasks": true,
  "SyncResources": true,
  "SyncAssignments": true,
  "AutoProvisionLists": true,
  "DeleteRemovedItems": true,
  "ShowSuccessNotification": true,
  "ShowErrorDialog": true,
  "ProjectsListTitle": "MSP Projects",
  "TasksListTitle": "MSP Tasks",
  "ResourcesListTitle": "MSP Resources",
  "AssignmentsListTitle": "MSP Assignments",
  "CustomTaskFields": [ "Text1", "Number1" ]
}
```

## Obmedzenia a poznámky

- Synchronizácia je **jednosmerná** (Project → SharePoint). Ručné úpravy v listoch sa pri ďalšej zmene
  danej úlohy prepíšu.
- MS Project nemá udalosť „po uložení“; dáta sa zachytia v `ProjectBeforeSave`. Ak v dialógu
  *Uložiť ako* stlačíte *Zrušiť*, synchronizácia aj tak prebehne s aktuálnymi dátami.
- Externé (vložené) úlohy z iných projektov sa preskakujú.
- Stĺpce jedného riadku textu sa orezávajú na 255 znakov, viacriadkové na 60 000.
- Ak sa prihlasovacie okno zruší pri automatickej synchronizácii, ďalšie automatické výzvy sa do
  najbližšieho *Sync now* nezobrazujú.
- Office.js (web add-in) pre Project nepodporuje udalosť uloženia ani prístup k celému projektu –
  preto VSTO.
