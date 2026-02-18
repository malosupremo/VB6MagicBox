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
