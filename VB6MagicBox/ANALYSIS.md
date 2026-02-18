# VB6MagicBox - Analisi e Modifiche

## ✅ PROBLEMI RISOLTI

### 1. Riferimenti Duplicati nei Campi delle Strutture
Le funzioni di risoluzione scansionavano tutto il file invece delle sole righe della procedura corrente.
```csharp
var startIndex = Math.Max(0, proc.StartLine - 1);
var endIndex = Math.Min(fileLines.Length, proc.EndLine);
for (int i = startIndex; i < endIndex; i++)
```

### 2. IndexOutOfRangeException
`StartLine`/`EndLine` non impostati causavano crash. Aggiunti controlli di sicurezza in tutte le funzioni di risoluzione.

### 3. Logica Errata per Determinare la Procedura Corrente
```csharp
// PRIMA: Procedures.FirstOrDefault(p => lineNum >= p.LineNumber)
// DOPO:  GetProcedureAtLine(lineNum) // usa StartLine/EndLine
```

### 4. Warning per Procedure API Esterne
`Declare Function`/`Declare Sub` senza corpo avevano `StartLine`/`EndLine` = 0. Soluzione: impostare entrambe alla riga di dichiarazione.

### 5. Accesso ai Campi degli Array Non Rilevato
`marrQueuePolling(i).Msg_h` non veniva rinominato. Regex `ReFieldAccess` migliorata per gestire le parentesi; aggiunta estrazione del nome base prima di `(`.

### 6. Property Get/Let con Nomi Duplicati
`PollingEnableRequest2` (suffisso errato). Se esiste già una Property con lo stesso nome base, riusa il suo `ConventionalName`.

### 7. Variabili Globali Non Idempotenti
`g_OpticalMonitor` → `GOpticalMonitor` alla seconda esecuzione. Aggiunta `IsPascalCase()`: se il tail dopo `g_` è già PascalCase, mantieni invariato.

### 8. Array di Controlli VB6
Ogni istanza `Begin VB.TextBox tbPower` creava un duplicato. Raggruppamento post-parsing per nome; aggiunta `LineNumbers` per tracciare tutte le posizioni.

### 9. Referenze Cross-Module per Controlli
`frmSQM242.tbPower(i).Text` non marcava il controllo come `Used`. Regex avanzata `ModuleName.ControlName(index).Property` con ricerca cross-module.

### 10. Parametri Declare Function su Più Righe
Solo il primo parametro veniva rinominato. `AddParameterReferencesForMultilineDeclaration()` segue i `_` e aggiunge References riga per riga.

### 11. Properties di Classe Trattate come Procedure Globali
**A — Modello separato** (`Models.cs`, `Parser.Core.cs`): `VbProperty` dedicata → `mod.Properties`, non `mod.Procedures`.
**B — References accumulate** (`Parser.Resolve.cs`): `alreadyInCalls` blocca solo duplicati nelle Calls; References accumulate per ogni riga.
**C — Dot-prefixed rename** (`Refactoring.cs`): fuori dalla classe usa `\.OldName\b` per non toccare parametri omonimi.

### 12. Rename Controlli Tocca il Nome della Classe
`Begin S7DATALib.S7Data S7Data` → lookbehind per rinominare solo il nome DOPO il secondo token:
```csharp
pattern = $@"(?<=^.*Begin\s+\S+\s+){Regex.Escape(oldName)}\b";
```

### 13. Tipi Usati Come `As TypeName` Non Rinominati
`ResolveTypeReferences()` aggiunge References al tipo per ogni `As TypeName` in: campi di Type, variabili globali/locali, parametri di procedure e proprietà.

### 14. Modulo `Used = false` con Membri Usati
Un modulo con costanti o enum usati restava `Used = false`. Fix: al termine di `BuildDependenciesAndUsage`, propagazione esplicita:
```csharp
mod.Used = mod.Procedures.Any(p => p.Used) || mod.Constants.Any(c => c.Used)
        || mod.Enums.Any(e => e.Used || e.Values.Any(v => v.Used)) || ...
```

### 15. Grafo Mermaid Incompleto (solo chiamate dirette)
`ExportMermaid` usava `project.Dependencies` (sole chiamate risolte). Sostituito con `ModuleReferences` (aggregato di tutte le References dei membri). Il grafo mostra ora anche dipendenze da costanti, tipi, enum, controlli. I nomi nel grafo usano `ConventionalName`.

### 16. `As New ClassName` — Tipo Parsato come `New`
`ReGlobalVar`, `ReLocalVar`, `ReMemberVar` catturavano `New` come tipo per `Private x As New TextFile`.
Fix: aggiunto `(?:New\s+)?` dopo `As\s+` in tutti e tre i regex. Conseguenza: `env["m_LogError"] = "TextFile"` anziché `"New"` → risoluzione dot-access corretta → `WriteLine` ottiene le sue References.

### 17. Classi Usate come Tipo Non Tracciate nelle References
`Private m_Log As New TextFile` non aggiungeva References alla classe `TextFile`.
Fix: `ResolveClassModuleReferences()` (modellata su `ResolveTypeReferences`) scansiona variabili globali, locali e parametri; se il tipo è una classe, aggiunge Reference al modulo classe.

### 18. Prefisso `Cls` Mancante per Moduli Classe
`TextFile.cls` → `ConventionalName = "TextFile"` (già PascalCase, nessun rename).
Fix: i file `.cls` ricevono prefisso `Cls` se non già presente, simmetrico al prefisso `frm` per i form.
```
TextFile.cls  → ClsTextFile
clsLog.cls    → ClsLog   (già prefissato)
```

### 19. Pipeline Analisi/Refactoring Duplicata e Incompleta
`RunRefactoringInteractive` non generava il Mermaid; `SortProject` veniva chiamato due volte.
Fix: helper condiviso `ExportProjectFiles(project, vbpPath)` che scrive sempre i 4 file di output. Ordine garantito: **ParseAndResolve → ExportProjectFiles → ApplyRenames**.

### 20. LineNumber Errato nelle References di Metodi Classe (PASS 1.5)
Sub VB6 chiamate senza parentesi (`m_LogError.CloseFile`) venivano catturate da **PASS 1.5** (`obj.method` senza `(`). I `VbCall` erano aggiunti senza `LineNumber`; PASS 1.5b (che lo impostava) veniva saltato per `alreadyInCalls`. In `BuildDependenciesAndUsage` il fallback produceva sempre `proc.LineNumber` (es. 1537 × 3).
Fix: aggiunto `LineNumber = li + 1` ai due `VbCall` in PASS 1.5 (property e procedure).

---

## ⚙️ MODELLO

```csharp
// VbProcedure
int StartLine, EndLine;
bool ContainsLine(int lineNumber);

// VbProperty (separata da VbProcedure)
string Name, ConventionalName, Kind, Visibility, ReturnType;
int StartLine, EndLine;
List<VbParameter> Parameters;
List<VbReference> References;
bool ContainsLine(int lineNumber);

// VbModule
List<VbProperty> Properties;       // separata da Procedures
List<string> ModuleReferences;     // moduli che referenziano questo (aggregato)
VbProject Owner;
bool IsForm, IsClass;
VbProcedure? GetProcedureAtLine(int lineNumber);

// VbControl
List<int> LineNumbers;             // per array di controlli
```

---

## 🔄 ORDINE DI ESECUZIONE

### FASE 1 - Parsing (`Parser.Core.cs`)
1. Function/Sub → `mod.Procedures`
2. Property Get/Let/Set → `mod.Properties` (separato!)
3. `StartLine` = riga di inizio, `EndLine` = riga `End ...`

### FASE 2 - Risoluzione (`Parser.Resolve.cs`)
1. `procIndex`: solo Procedures (no Properties)
2. Accessi con punto — 3 pass nell'inner loop:
   - **PASS 1** `obj.method(` — con parentesi; risolve classe o procIndex
   - **PASS 1.5** `obj.method` senza `(` — Sub senza parens; risolve classe
   - **PASS 1.5b** `obj.method` generico — saltato se già in Calls
3. References accumulate per ogni occorrenza
4. `ResolveTypeReferences()`: References ai VbTypeDef per ogni `As TypeName`
5. `ResolveClassModuleReferences()`: References ai moduli classe per ogni `As [New] ClassName`
6. `MarkUsedTypes()`: propaga `Used` ai tipi
7. Propagazione `Used` ai moduli: se qualunque membro è usato, il modulo è usato
8. Costruzione `ModuleReferences`: aggregato dei caller da tutti i membri

### FASE 3 - Export (`Parser.Export.cs`)
`ExportProjectFiles` scrive sempre: `symbols.json`, `rename.json`, `rename.csv`, `dependencies.md`

### FASE 4 - Refactoring (`Refactoring.cs`)

| Contesto                    | Pattern                            |
|-----------------------------|------------------------------------|
| Controllo su riga `Begin`   | `(?<=Begin\s+\S+\s+)OldName\b`     |
| Property fuori dalla classe | `\.OldName\b` → `.NewName`         |
| Tutto il resto              | `\bOldName\b`                      |

Ordine applicazione: Field → EnumValue → Type → Enum → Constant → GlobalVariable → PropertyParameter → Parameter → LocalVariable → Control → Property → Procedure → Module

---

## 📋 TODO
- 🔧 Prefissi controlli: migliorare mappa dei prefissi standard VB6
- 🔧 Verifica compilazione del codice refactorizzato
- 🗑️ Rimozione/commento variabili non usate
- 📝 Aggiunta `As NomeTipo` dove mancante
- 🪄 Comando MAGIC: esegue tutte le modifiche in automatico

---

## 🔧 NOTE TECNICHE

| Voce             | Valore                            |
|------------------|-----------------------------------|
| Target Framework | .NET 10                           |
| Encoding         | Windows-1252 (VB6)                |
| Idempotenza      | verificata (Gennaio 2025)         |
| Test VBP         | caller.vbp + progetto produzione  |


## ✅ PROBLEMI RISOLTI

### 1. Riferimenti Duplicati nei Campi delle Strutture
Le funzioni di risoluzione scansionavano tutto il file invece delle sole righe della procedura corrente.
```csharp
var startIndex = Math.Max(0, proc.StartLine - 1);
var endIndex = Math.Min(fileLines.Length, proc.EndLine);
for (int i = startIndex; i < endIndex; i++)
```

### 2. IndexOutOfRangeException
`StartLine`/`EndLine` non impostati causavano crash. Aggiunti controlli di sicurezza in tutte le funzioni di risoluzione.

### 3. Logica Errata per Determinare la Procedura Corrente
```csharp
// PRIMA: Procedures.FirstOrDefault(p => lineNum >= p.LineNumber)
// DOPO:  GetProcedureAtLine(lineNum) // usa StartLine/EndLine
```

### 4. Warning per Procedure API Esterne
`Declare Function`/`Declare Sub` senza corpo avevano `StartLine`/`EndLine` = 0. Soluzione: impostare entrambe alla riga di dichiarazione.

### 5. Accesso ai Campi degli Array Non Rilevato
`marrQueuePolling(i).Msg_h` non veniva rinominato. Regex `ReFieldAccess` migliorata per gestire le parentesi; aggiunta estrazione del nome base prima di `(`.

### 6. Property Get/Let con Nomi Duplicati
`PollingEnableRequest2` (suffisso errato). Se esiste già una Property con lo stesso nome base, riusa il suo `ConventionalName`.

### 7. Variabili Globali Non Idempotenti
`g_OpticalMonitor` → `GOpticalMonitor` alla seconda esecuzione. Aggiunta `IsPascalCase()`: se il tail dopo `g_` è già PascalCase, mantieni invariato.

### 8. Array di Controlli VB6
Ogni istanza `Begin VB.TextBox tbPower` creava un duplicato. Raggruppamento post-parsing per nome; aggiunta `LineNumbers` per tracciare tutte le posizioni.

### 9. Referenze Cross-Module per Controlli
`frmSQM242.tbPower(i).Text` non marcava il controllo come `Used`. Regex avanzata `ModuleName.ControlName(index).Property` con ricerca cross-module.

### 10. Parametri Declare Function su Più Righe
Solo il primo parametro veniva rinominato. `AddParameterReferencesForMultilineDeclaration()` segue i `_` e aggiunge References riga per riga.

### 11. Properties di Classe Trattate come Procedure Globali
**3 cause, 3 fix:**

**A — Modello separato** (`Models.cs`, `Parser.Core.cs`): classe `VbProperty` dedicata, aggiunta a `mod.Properties`, **non** a `mod.Procedures`.

**B — References accumulate** (`Parser.Resolve.cs`): il vecchio `Skip if already in calls` perdeva le occorrenze successive. Ora `alreadyInCalls` blocca solo i duplicati nelle Calls; le References vengono sempre accumulate per ogni riga.

**C — Dot-prefixed rename** (`Refactoring.cs`): fuori dalla classe usa `\.OldName\b` per non toccare parametri con lo stesso nome.

```
PRIMA: g_PlasmaSource.IsDeposit = IsDeposit   <- entrambi rinominati
DOPO:  g_PlasmaSource.IsDeposit = isDeposit   <- solo .IsDeposit rinominato ✅
```

### 12. Rename Controlli Tocca il Nome della Classe
`Begin S7DATALib.S7Data S7Data` diventava `Begin S7DATALib.sdtS7Data sdtS7Data` (anche il tipo veniva rinominato).
Lookbehind per sostituire solo il nome DOPO il secondo token:
```csharp
pattern = $@"(?<=^.*Begin\s+\S+\s+){Regex.Escape(oldName)}\b";
```

### 13. Tipi Usati Come `As TypeName` Non Rinominati
`DspH As DISPAT_HEADER_T` non veniva rinominato perché il tipo non aveva References su quella riga.
Causa: il type index tracciava solo la dichiarazione del tipo, non i punti di utilizzo.
Soluzione: `ResolveTypeReferences()` aggiunge References al tipo per ogni occorrenza in:
- Campi di altri Type: `FieldName As OTHER_TYPE`
- Variabili globali: `Dim x As MY_TYPE`
- Parametri di procedure e proprietà
- Variabili locali
```csharp
// Chiamato alla fine di ResolveTypesAndCalls()
private static void ResolveTypeReferences(VbProject project, Dictionary<string, VbTypeDef> typeIndex)
// Helper:
private static void AddTypeReference(string typeName, int lineNumber,
    string moduleName, string procedureName, Dictionary<string, VbTypeDef> typeIndex)
```

---

## ⚙️ MODELLO

```csharp
// VbProcedure
int StartLine, EndLine;
bool ContainsLine(int lineNumber);

// VbProperty (NUOVA - separata da VbProcedure)
string Name, ConventionalName, Kind, Visibility, ReturnType;
int StartLine, EndLine;
List<VbParameter> Parameters;
List<VbReference> References;
bool ContainsLine(int lineNumber);

// VbModule
List<VbProperty> Properties;       // separata da Procedures
VbProject Owner;
bool IsForm, IsClass;
VbProcedure? GetProcedureAtLine(int lineNumber);

// VbControl
List<int> LineNumbers;             // per array di controlli
```

---

## 🔄 ORDINE DI ESECUZIONE

### FASE 1 - Parsing (`Parser.Core.cs`)
1. Function/Sub → `mod.Procedures`
2. Property Get/Let/Set → `mod.Properties` (separato!)
3. `StartLine` = riga di inizio, `EndLine` = riga `End ...`

### FASE 2 - Risoluzione (`Parser.Resolve.cs`)
1. `procIndex`: solo Procedures (no Properties)
2. Accessi con punto: cerca prima in Properties, poi in Procedures
3. References: accumulate per ogni occorrenza (no skip per Properties)
4. `ResolveTypeReferences()`: References ai tipi per ogni `As TypeName`

### FASE 3 - Refactoring (`Refactoring.cs`)

| Contesto                   | Pattern                            |
|----------------------------|------------------------------------|
| Controllo su riga `Begin`  | `(?<=Begin\s+\S+\s+)OldName\b`     |
| Property fuori dalla classe| `\.OldName\b` → `.NewName`          |
| Tutto il resto             | `\bOldName\b`                      |

---

## 📋 TODO
- ✅ Array Controlli, Idempotenza, Properties Separate, Type References, Rename Begin
- 🔧 Prefissi controlli: migliorare mappa dei prefissi standard VB6
- 🔧 Verifica compilazione del codice refactorizzato
- 🗑️ Rimozione/commento variabili non usate
- 📝 Aggiunta `As NomeTipo` dove mancante
- 🪄 Comando MAGIC: esegue tutte le modifiche in automatico

---

## 🔧 NOTE TECNICHE

| Voce             | Valore                            |
|------------------|-----------------------------------|
| Target Framework | .NET 10                           |
| Encoding         | Windows-1252 (VB6)                |
| Idempotenza      | verificata (Gennaio 2025)         |
| Test VBP         | caller.vbp + progetto produzione  |
