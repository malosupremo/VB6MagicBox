# VB6MagicBox — Naming Conventions

Documento generato dall'analisi di `Parser.Naming.cs` e `Refactoring.cs`.  
Ogni sezione riporta il tipo di simbolo, la regola applicata, esempi concreti e note speciali.

---

## Indice

1. [Moduli / File](#1-moduli--file)
2. [Variabili Globali](#2-variabili-globali)
3. [Costanti (livello modulo)](#3-costanti-livello-modulo)
4. [Enum](#4-enum)
5. [Valori Enum](#5-valori-enum)
6. [Tipi UDT (`Type…End Type`)](#6-tipi-udt-typeend-type)
7. [Campi dei Tipi UDT](#7-campi-dei-tipi-udt)
8. [Controlli Form](#8-controlli-form)
9. [Procedure e Funzioni](#9-procedure-e-funzioni)
10. [Property (Get / Let / Set)](#10-property-get--let--set)
11. [Parametri di Procedure e Property](#11-parametri-di-procedure-e-property)
12. [Variabili Locali](#12-variabili-locali)
13. [Costanti Locali](#13-costanti-locali)
14. [Parametri di Evento (`Event`)](#14-parametri-di-evento-event)
15. [Rimozione Prefissi Ungheresi](#15-rimozione-prefissi-ungheresi)
16. [Conflict Resolution](#16-conflict-resolution)
17. [Ordine di Applicazione dei Rename](#17-ordine-di-applicazione-dei-rename)
18. [Pattern Regex per il Rename](#18-pattern-regex-per-il-rename)
19. [Casi Speciali](#19-casi-speciali)

---
## Riassunto

| Sezione |	Contenuto |
| ------- | --------- | 
| §1 Moduli	 | .bas/.cls/.frm, prefisso `Frm`/`Cls` (sempre maiuscolo), all-caps+cifre → PascalCase per run di lettere, keyword → Class|
| §2 Variabili Globali	| 7 sotto-casi: private, static, gobj, classe in .bas, classe in form, primitivo, suffisso _N |
| §3 Costanti modulo |	ToScreamingSnakeCase(string) con esempi di acronimi |
| §4-5 Enum / Valori |	ToPascalCaseFromScreamingSnake(string) |
| §6 Tipi UDT |	ToPascalCaseType(string) con suffisso _T sempre ri-aggiunto |
| §7 Campi UDT |	PascalCase, Type → TypeValue |
| §8 Controlli |	Tabella completa dei 30+ prefissi, logica di sostituzione, tb/tx → txt |
| §9-10 Procedure / Property |	Event handler standard, controllo, WithEvents, normali; regola Get/Let/Set condivisi |
| §11-14 Parametri / Locali / Costanti locali |	camelCase, s_, strip ungherese |
| §15 Prefissi ungheresi |	Tabella prefissi rimossi + preservati (msg, plc) + casi speciali m, m_, cur |
| §16 Conflict resolution |	Suffisso numerico, fallback a nome originale per keyword |
| §17 Ordine rename |	Tabella priorità 1-13 con motivazione |
| §18 Pattern regex |	3 pattern distinti per contesto |
| §19 Casi speciali |	Attribute VB_Name, encoding Windows-1252, backup, idempotenza |


---

## 1. Moduli / File

**Regola:** `ToPascalCase` sul nome base del file (senza estensione e senza prefisso).
Form e classi ricevono un prefisso fisso (`Frm` / `Cls`) **prima** di applicare la conversione al nome base.

| Tipo file | Regola | Esempio prima | Esempio dopo |
|-----------|--------|---------------|--------------|
| `.bas` modulo | PascalCase | `GLOBAL_MODULE` | `GlobalModule` |
| `.bas` modulo (all-caps con cifre) | PascalCase per run di lettere | `EXEC1` | `Exec1` |
| `.bas` modulo (all-caps + cifre misti) | PascalCase per run di lettere | `SQM242HND` | `Sqm242Hnd` |
| `.cls` classe | `Cls` + PascalCase | `TextFile` | `ClsTextFile` |
| `.cls` classe (già prefissata) | `Cls` + PascalCase del nome base | `clsTextFile` | `ClsTextFile` |
| `.cls` classe (all-caps) | `Cls` + PascalCase | `LOGGER` | `ClsLogger` |
| `.frm` form | `Frm` + PascalCase del nome base | `Main_Form` | `FrmMainForm` |
| `.frm` form (all-caps con cifre) | `Frm` + PascalCase del nome base | `SQM242` | `FrmSqm242` |
| `.frm` form (già prefissata, all-caps) | `Frm` + PascalCase del nome base | `frmSQM242` | `FrmSqm242` |

**Algoritmo `ToPascalCase` per segmenti all-uppercase:**

Un segmento (parte separata da `_`) viene convertito in PascalCase se **tutte le sue lettere sono maiuscole** (le cifre non partecipano al controllo). La conversione applica una regex che processa ogni _run di lettere_ separatamente, lasciando le cifre invariate nella loro posizione:

```

---

## 20. Output diagnostici (CSV)

| File | Contenuto | Note |
|------|-----------|------|
| `*.disambiguations.csv` | Solo righe dove è stato applicato un prefisso di disambiguazione | Ordinato per `Module` e `LineNumber` |
| `*.shadows.csv` | Conflitti locali vs simboli esterni | Include `LineNumber`, `LocalType`, `ShadowedType` |

---

## 21. Spacing rules (formatter)

| Regola | Dettaglio |
|--------|-----------|
| Property Get/Let/Set | Blocchi con stesso nome restano adiacenti, senza riga vuota |
| Post dichiarazioni | Dopo l’ultimo `Dim/Static/Const` va sempre una riga vuota |
| Prima dei loop | `For/Do` hanno riga vuota se preceduti da istruzioni non‑blocco |
| `End With` | Inserisce riga vuota se non seguito da un altro `End ...` |
"SQM242HND" → run "SQM" → "Sqm", cifre "242" invariate, run "HND" → "Hnd" → "Sqm242Hnd"
"EXEC1"     → run "EXEC" → "Exec", cifra "1" invariata                → "Exec1"
"MAX"       → run "MAX" → "Max"                                        → "Max"
```

I segmenti con lettere miste (es. `frmMain`, `PDxI`) seguono la regola precedente: prima lettera maiuscola, resto invariato.

**Note:**
- Il prefisso `Frm`/`Cls` viene aggiunto **dopo** la PascalCase del nome base (non prima), così `SQM242.frm` → `FrmSqm242` e non `FrmSQM242`.
- Se il nome risultante è una keyword C# (es. `With`), viene aggiunto il suffisso `Class` → `WithClass`.
- Nelle classi/form VB6 viene aggiornata anche la riga `Attribute VB_Name = "..."`.

---

## 2. Variabili Globali

La regola dipende da **visibilità**, **staticità**, **tipo** e **tipo di modulo** (`.bas` vs `.frm`/`.cls`).

### 2a. Private (`Dim` / `Private`)

| Caso | Regola | Esempio prima | Esempio dopo |
|------|--------|---------------|--------------|
| Già formato `m_Nome` | Invariato | `m_Counter` | `m_Counter` |
| Tutto il resto | `m_` + strip prefisso ungherese + PascalCase | `mudtPollingState` | `m_PollingState` |

### 2b. Static

| Caso | Regola | Esempio prima | Esempio dopo |
|------|--------|---------------|--------------|
| Già formato `s_Nome` | Invariato | `s_Retry` | `s_Retry` |
| Prefisso `s` + maiuscola (`sCounter`) | `s_` + PascalCase sul resto | `sCounter` | `s_Counter` |
| Tutto il resto | `s_` + PascalCase | `static_val` | `s_StaticVal` |

### 2c. Public / Global — oggetti con prefisso `gobj`

| Regola | Esempio prima | Esempio dopo |
|--------|---------------|--------------|
| Rimuovi `gobj`, aggiungi `g_` + PascalCase | `gobjPlcManager` | `g_PlcManager` |

### 2d. Public / Global — tipo classe, in modulo `.bas`

| Caso | Regola | Esempio prima | Esempio dopo |
|------|--------|---------------|--------------|
| Già `g_` + PascalCase tail | Invariato | `g_OpticalMonitor` | `g_OpticalMonitor` |
| Già `g_` + tail non-PascalCase | `g_` + PascalCase del tail | `g_optical_monitor` | `g_OpticalMonitor` |
| Nessun prefisso | `g_` + PascalCase | `OpticalMonitor` | `g_OpticalMonitor` |

### 2e. Public / Global — tipo classe, in form (`.frm`)

| Regola | Esempio prima | Esempio dopo |
|--------|---------------|--------------|
| Rimuovi eventuale `obj`, aggiungi `obj` + PascalCase | `UAServerObj` | `objUaServerObj` |
| Già `obj` | `obj` + PascalCase del resto | `objFM489` | `objFm489` |

### 2f. Public / Global — tipo classe, in classe (`.cls`)

| Regola | Esempio prima | Esempio dopo |
|--------|---------------|--------------|
| PascalCase puro (come properties) — strip prefisso `obj` se presente | `ServerObj` | `ServerObj` |
| | `objUAServerObj` | `UAServerObj` |

### 2f. Public / Global — tipo primitivo (non classe)

| Regola | Esempio prima | Esempio dopo |
|--------|---------------|--------------|
| Strip prefisso ungherese + PascalCase | `intPollingCount` | `PollingCount` |

### 2g. Suffisso array `_N`

Il suffisso numerico (es. `_0`, `_1`) viene preservato dopo la conversione del nome base.

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `objFM489_0` | `objFm489_0` |
| `g_Connection_1` | `g_Connection_1` |

---

## 3. Costanti (livello modulo)

**Regola:** SCREAMING_SNAKE_CASE via `ToScreamingSnakeCase()`.

La funzione inserisce `_` alle transizioni:
- lettera minuscola → MAIUSCOLA (`itemUA` → `ITEM_UA`)
- fine di acronimo prima di parola minuscola (`UAObj` → `UA_OBJ`)
- separa già sugli `_` esistenti

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `ItemUAObjListener` | `ITEM_UA_OBJ_LISTENER` |
| `MaxUnsignedLongAnd1` | `MAX_UNSIGNED_LONG_AND1` |
| `RIC_AL_PDxI1_LOWVOLTAGE` | `RIC_AL_PDXI1_LOWVOLTAGE` *(underscore preservati)* |
| `Alg_FirstStep` | `ALG_FIRST_STEP` |

**Nota:** le costanti già in `SCREAMING_SNAKE_CASE` (`^[A-Z0-9_]+$`) vengono preservate invariate da `ToUpperSnakeCase`.

---

## 4. Enum

**Regola:** PascalCase da SCREAMING_SNAKE_CASE via `ToPascalCaseFromScreamingSnake()`.  
Ogni segmento separato da `_` diventa Parola con iniziale maiuscola e resto minuscolo.

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `POLLING_STATE` | `PollingState` |
| `UA_SERVER_STATUS` | `UaServerStatus` |
| `AlreadyCamelCase` | `AlreadyCamelcase` *(nessun `_`, trattato come PascalCase)* |

---

## 5. Valori Enum

**Regola:** stessa di [§4](#4-enum) — PascalCase da SCREAMING_SNAKE_CASE.

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `POLL_ACTIVE` | `PollActive` |
| `STATE_IDLE_WAIT` | `StateIdleWait` |

---

## 6. Tipi UDT (`Type…End Type`)

**Regola:** PascalCase via `ToPascalCaseType()`, con suffisso `_T` sempre ri-aggiunto.

Algoritmo:
1. Rimuovi eventuale `_T` o `_T` finale (`HEADER_T` → `HEADER`, `HEADERT` non tocca).
2. Dividi per `_`.
3. Ogni parte: se tutta maiuscola → PascalCase (prima maiuscola, resto minuscolo); altrimenti prima maiuscola + resto invariato.
4. Ri-aggiungi `_T`.

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `DISPAT_HEADER_T` | `DispatHeader_T` |
| `POLLING_DATA_T` | `PollingData_T` |
| `MyTypeT` | `MyTypeT_T` *(raro, già senza `_T`)* |

---

## 7. Campi dei Tipi UDT

**Regola:** PascalCase via `ToPascalCase()`.

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `msg_data` | `MsgData` |
| `Msg_h` | `MsgH` |
| `Type` *(keyword)* | `TypeValue` |
| qualsiasi keyword C# | `<NomeOrigine>Value` |

Conflict resolution interno al tipo (scope isolato dal resto del modulo).

---

## 8. Controlli Form

**Regola:** prefisso standard da tipo + PascalCase del nome base.

### Prefissi standard

| Tipo VB6 | Prefisso | Tipo VB6 | Prefisso |
|----------|----------|----------|----------|
| `TextBox` | `txt` | `Label` | `lbl` |
| `CommandButton` / `Command` | `cmd` | `Frame` | `fra` |
| `CheckBox` | `chk` | `OptionButton` / `Option` | `opt` |
| `ListBox` | `lst` | `ComboBox` | `cbo` |
| `Timer` | `tmr` | `PictureBox` | `pic` |
| `Image` | `img` | `Shape` | `shp` |
| `Line` | `lin` | `HScrollBar` | `hsb` |
| `VScrollBar` | `vsb` | `DirListBox` | `dir` |
| `DriveListBox` | `drv` | `FileListBox` | `fil` |
| `Data` | `dat` | `OLE` | `ole` |
| `CommonDialog` | `dlg` | `Menu` | `mnu` |
| `MSFlexGrid` / `MSHFlexGrid` | `flx` | `DataGrid` | `grd` |
| `TreeView` | `tvw` | `ListView` | `lvw` |
| `ProgressBar` | `prg` | `Slider` | `sld` |
| `TabStrip` | `tab` | `ToolBar` | `tlb` |
| `StatusBar` | `stb` | `ImageList` | `iml` |
| `RichTextBox` | `rtf` | `MonthView` | `mvw` |
| `DateTimePicker` | `dtp` | `UpDown` | `upd` |
| `Animation` | `ani` | `MSComm` | `msc` |
| `Winsock` | `wsk` | `WebBrowser` | `web` |
| `CoolBar` | `clb` | `FlatScrollBar` | `fsb` |

**Tipo sconosciuto:** prefisso derivato: prima lettera + consonanti (fino a 3 caratteri) in lowercase.

### Logica di applicazione

| Caso | Regola | Esempio prima | Esempio dopo |
|------|--------|---------------|--------------|
| Già prefisso corretto + maiuscola dopo | Invariato | `txtNome` | `txtNome` |
| Prefisso `tb` o `tx` su TextBox | Normalizza a `txt` | `tbPower` | `txtPower` |
| Ha un altro prefisso a 3 lettere | Sostituisci prefisso | `lblCounter` (su ComboBox) | `cboCounter` |
| Nessun prefisso | Aggiungi prefisso + PascalCase | `Power` (TextBox) | `txtPower` |

**Namespace VB6 rimossi prima della ricerca:** `VB.`, `MSComCtl2.`, `MSComctlLib.`, `Threed.SS`.

**Array di controlli:** tutti gli elementi dello stesso nome condividono il medesimo `ConventionalName`.

---

## 9. Procedure e Funzioni

### 9a. Event handler standard — invariati

| Pattern | Esempio |
|---------|---------|
| `Class_Initialize` | `Class_Initialize` |
| `Class_Terminate` | `Class_Terminate` |
| `Form_*` (nei form) | `Form_Load`, `Form_Unload` |
| `UserControl_*` (nei form) | `UserControl_Initialize` |

### 9b. Event handler di controllo — `ControlConventionalName_EventPascalCase`

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `txtPower_Change` | `txtPower_Change` *(già corretto se il controllo lo è)* |
| `Command1_Click` | `cmd1_Click` *(usa ConventionalName del controllo)* |

### 9c. Event handler di variabile WithEvents — `VarConventionalName_EventPascalCase`

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `objATH3204_0_ReadStatusCompleted` | `objATH32040_ReadStatusCompleted` |

### 9d. Procedure normali — PascalCase

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `is_ready_to_start` | `IsReadyToStart` |
| `getPLCstatus` | `GetPLCstatus` |

---

## 10. Property (Get / Let / Set)

**Regola:** PascalCase — identica a [§9d](#9d-procedure-normali--pascalcase).

**Caso speciale:** tutte le varianti Get/Let/Set dello stesso nome base **condividono** lo stesso `ConventionalName` (la prima trovata determina il nome, le successive lo riutilizzano).

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `Property Get IsDeposit` | `Property Get IsDeposit` *(già PascalCase)* |
| `Property Let isDeposit` | `Property Let IsDeposit` *(allineato al Get)* |
| `Property Get polling_enable` | `Property Get PollingEnable` |
| `Property Let polling_enable` | `Property Let PollingEnable` *(stesso nome)* |

**Rename cross-module:** fuori dal modulo che definisce la property, il pattern usato è `\.OldName\b` → `.NewName` per evitare di toccare parametri o variabili omonimi.

```vb6
' PRIMA:
g_PlasmaSource.IsDeposit = IsDeposit   ' <- entrambi rinominati (BUG)

' DOPO:
g_PlasmaSource.IsDeposit = isDeposit   ' <- solo .IsDeposit rinominato ✅
```

---

## 11. Parametri di Procedure e Property

**Regola:** camelCase via `ToCamelCase()` (= PascalCase con prima lettera minuscola).

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `PollingEnabled` | `pollingEnabled` |
| `int_Value` | `intValue` |
| `strName` | `strName` *(nessun strip su params — no Hungarian stripping)* |

Conflict resolution in uno scope locale per ogni procedura/property.

---

## 12. Variabili Locali

### 12a. Static

| Caso | Regola | Esempio prima | Esempio dopo |
|------|--------|---------------|--------------|
| Già `s_Nome` | Invariato | `s_Retry` | `s_Retry` |
| Prefisso `s` + maiuscola | `s_` + PascalCase del resto | `sCounter` | `s_Counter` |
| Tutto il resto | `s_` + strip ungherese + PascalCase | `static_val` | `s_StaticVal` |

### 12b. Normali

**Regola:** strip prefisso ungherese + camelCase.

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `intCounter` | `counter` |
| `strMessage` | `message` |
| `bolEnabled` | `enabled` |
| `msgData` | `msgData` *(prefisso `msg` preservato)* |
| `plcStatus` | `plcStatus` *(prefisso `plc` preservato)* |

---

## 13. Costanti Locali

**Regola:** SCREAMING_SNAKE_CASE — stessa logica delle [costanti di modulo](#3-costanti-livello-modulo).

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `MaxRetry` | `MAX_RETRY` |
| `TimeoutMs` | `TIMEOUT_MS` |

---

## 14. Parametri di Evento (`Event`)

**Regola:** camelCase — identica ai [parametri di procedura](#11-parametri-di-procedure-e-property).

| Esempio prima | Esempio dopo |
|---------------|--------------|
| `StatusCode` | `statusCode` |
| `ErrorMessage` | `errorMessage` |

---

## 15. Rimozione Prefissi Ungheresi

Funzione `GetBaseNameFromHungarian()` — applicata a variabili globali e locali (non ai parametri).

### Prefissi rimossi

| Prefisso | Tipo implicato |
|----------|----------------|
| `int` | Integer |
| `str` | String |
| `lng` | Long |
| `dbl` | Double |
| `sng` | Single |
| `cur` | Currency *(solo se il tipo dichiarato è `Currency`)* |
| `bol` | Boolean |
| `byt` | Byte |
| `chr` | String/Char |
| `dat` | Date |
| `obj` | Object |
| `arr` | Array |
| `udt` | User-Defined Type |

### Prefissi preservati (NOT stripped)

| Prefisso | Motivo |
|----------|--------|
| `msg` | Non è un tipo ungherese standard |
| `plc` | Acronimo di dominio (Programmable Logic Controller) |

### Casi speciali

| Pattern | Comportamento | Esempio |
|---------|---------------|---------|
| `m` + lettera maiuscola | Rimuovi la `m` | `mLogger` → `Logger` |
| `m_Nome` | Già corretto, invariato | `m_Counter` → `m_Counter` |
| `s_Nome` | Già corretto, invariato | `s_Retry` → `s_Retry` |
| `mCCC...` (Hungarian `mudtXxx`) | Strip del pattern `m[a-z]{3}` | `mudtPollingState` → `PollingState` |
| `cur` con tipo non-Currency | Preservato (= "current") | `curPosition` → `curPosition` |

---

## 16. Conflict Resolution

Se il `ConventionalName` proposto è già occupato nello stesso scope, viene aggiunto un suffisso numerico progressivo.

| Proposto | Già usato | Risultato |
|----------|-----------|-----------|
| `result` | — | `result` |
| `result` | `result` | `result2` |
| `result` | `result`, `result2` | `result3` |

Se il nome finale è una **keyword C# riservata** (`abstract`, `class`, `string`, `With`, …) il simbolo torna al **nome originale VB6**.

---

## 17. Ordine di Applicazione dei Rename

I rename vengono applicati dal più specifico (scope locale) al più generale, per evitare sostituzioni accidentali.

| Priorità | Categoria | Motivo |
|----------|-----------|--------|
| 1 | `Field` | Membro di Type — prima del nome del Type |
| 2 | `EnumValue` | Valore enum — prima del nome dell'Enum |
| 3 | `Type` | Dipende dai Field già rinominati |
| 4 | `Enum` | Dipende dai Values già rinominati |
| 5 | `Constant` | Nessuna dipendenza, ma molti la usano |
| 6 | `GlobalVariable` | Può istanziare un Type |
| 7 | `PropertyParameter` | Scope locale a Property |
| 8 | `Parameter` | Scope locale a Procedure |
| 9 | `LocalVariable` | Scope locale a Procedure |
| 10 | `Control` | Form-specific |
| 11 | `Property` | Accessi con punto, cross-module |
| 12 | `Procedure` | Visibili globalmente |
| 13 | `Module` | Top-level, meno specifico |

A parità di priorità, i nomi più **lunghi** vengono rinominati per primi (per evitare sostituzioni parziali di sottostringhe).

---

## 18. Pattern Regex per il Rename

Il rename viene applicato **solo sulle righe elencate nei `LineNumbers` delle References** (dati della Fase 1). Per ogni riga il commento VB6 (`'...`) è preservato: la sostituzione avviene solo nella parte di codice.

| Contesto | Pattern usato | Sostituzione |
|----------|---------------|--------------|
| Riga `Begin Library.Type ControlName` | `(?<=^.*Begin\s+\S+\s+)OldName\b` | `NewName` |
| Property **fuori** dal modulo che la definisce | `\.OldName\b` | `.NewName` |
| Tutto il resto | `\bOldName\b` | `NewName` |

---

## 19. Casi Speciali

### Attributi VB6

| Riga | Aggiornata quando |
|------|-------------------|
| `Attribute VB_Name = "ClassName"` | Rename del modulo (classi e form) |
| `Attribute OldName.VB_VarHelpID = -1` | Rename di una variabile globale (riga successiva alla dichiarazione) |

### Encode file

I file VB6 vengono letti e riscritti con encoding **Windows-1252** (ANSI), richiesto da `.NET Core`/`.NET 5+` tramite `Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)`.

### Backup progressivo

Prima di modificare un file viene creato un backup nella cartella `<ParentDir>.backup<yyyyMMdd_HHmmss>`. Solo i file effettivamente modificati vengono copiati nel backup.

### Idempotenza

Le convenzioni sono progettate per essere **idempotenti**: eseguire il tool due volte sullo stesso sorgente non produce ulteriori modifiche. Verificato tramite:
- `IsPascalCase()`: se il tail dopo `g_` è già PascalCase, il nome viene mantenuto invariato.
- `IsConventional` su ogni simbolo: se `Name == ConventionalName` il simbolo viene saltato.
