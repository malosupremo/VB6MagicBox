# VB6MagicBox - Analisi e Modifiche Applicate

## ?? **PROBLEMI RISOLTI**

### 1. **Riferimenti Duplicati nei Campi delle Strutture**
**PROBLEMA**: Il campo `Msg_h` appariva con gli stessi numeri di riga (218, 232, 236, 243) in procedure diverse come `Add`, `Class_Initialize`, `CompactArray`, `Count`.

**CAUSA RADICE**: Le funzioni di risoluzione scansionavano tutto il file dalla procedura corrente fino alla fine, invece che solo le righe della procedura corrente.

**FUNZIONI COINVOLTE**:
- `ResolveFieldAccesses()` - Accessi ai campi delle strutture (var.field)
- `ResolveControlAccesses()` - Accessi ai controlli dei form
- `ResolveParameterAndLocalVariableReferences()` - Variabili locali e parametri
- Loop principale per rilevamento procedure nude

**SOLUZIONE IMPLEMENTATA**:
```csharp
// PRIMA (SBAGLIATO):
for (int i = proc.LineNumber - 1; i < fileLines.Length; i++)

// DOPO (CORRETTO):
var startIndex = Math.Max(0, proc.StartLine - 1);
var endIndex = Math.Min(fileLines.Length, proc.EndLine);
for (int i = startIndex; i < endIndex; i++)
```

### 2. **IndexOutOfRangeException**
**PROBLEMA**: Crash dell'applicazione durante la risoluzione dei riferimenti.

**CAUSA**: `StartLine`/`EndLine` non impostati o con valori non validi.

**SOLUZIONE**: Controlli di sicurezza in tutte le funzioni di risoluzione:
```csharp
if (proc.StartLine <= 0) proc.StartLine = proc.LineNumber;
if (proc.EndLine <= 0) proc.EndLine = fileLines.Length;
var startIndex = Math.Max(0, proc.StartLine - 1);
var endIndex = Math.Min(fileLines.Length, proc.EndLine);
```

### 3. **Logica Errata per Determinare la Procedura Corrente**
**PROBLEMA**: Variabili globali assegnate erroneamente alle procedure.

**PRIMA**:
```csharp
var procAtLine = searchMod.Procedures.FirstOrDefault(p => lineNum >= p.LineNumber);
```

**DOPO**:
```csharp
var procAtLine = searchMod.GetProcedureAtLine(lineNum); // Usa StartLine/EndLine
```

### 4. **Warning per Procedure API Esterne**
**PROBLEMA**: Warning di tipo `Procedure Sleep has invalid EndLine: 0` per le procedure `Declare Function`/`Declare Sub`.

**CAUSA**: Le procedure API esterne sono dichiarate su una sola riga e non hanno corpo, quindi `StartLine`/`EndLine` non venivano impostate.

**SOLUZIONE**: Impostare sia `StartLine` che `EndLine` alla stessa riga per procedure API esterne:
```csharp
// Declare Function/Sub: StartLine = EndLine = originalLineNumber
StartLine = originalLineNumber,
EndLine = originalLineNumber
```

### 5. **Accesso ai Campi degli Array Non Rilevato**
**PROBLEMA**: I campi delle strutture accessibili tramite array non venivano rilevati per il refactoring.
- ✅ `Item.Msg_h` → `Item.MsgH` (funzionava)
- ❌ `marrQueuePolling(i).Msg_h` → rimaneva `Msg_h` (non funzionava)

**CAUSA RADICE**: La regex `ReFieldAccess` non gestiva l'accesso agli array con parentesi.

**SOLUZIONE**:
1. **Regex migliorata**: `(\w+)\.(\w+)` → `([A-Za-z_]\w*(?:\([^)]*\))?)\s*\.\s*([A-Za-z_]\w+)`
2. **Estrazione nome base**: `marrQueuePolling(i)` → `marrQueuePolling` per cercare nell'ambiente

```csharp
// Estrai il nome base della variabile rimuovendo l'accesso array
var baseVarName = varName;
var parenIndex = varName.IndexOf('(');
if (parenIndex >= 0)
{
  baseVarName = varName.Substring(0, parenIndex);
}
```

### 6. **Property Get/Let con Nomi Duplicati**
**PROBLEMA**: Le Property Get e Let ricevevano nomi diversi per evitare conflitti.
```csharp
Property Get PollingEnableRequest() // OK
Property Let PollingEnableRequest2() // ❌ Suffisso "2" errato
```

**CAUSA**: Il sistema di conflict resolution non distingueva le Property Get/Let/Set come appartenenti alla stessa proprietà.

**SOLUZIONE**: Raggrupamento delle Property per nome base:
```csharp
// SPECIALE: Property Get/Let/Set con lo stesso nome devono mantenere lo stesso ConventionalName
if (proc.Kind.StartsWith("Property", StringComparison.OrdinalIgnoreCase))
{
  // Cerca se esiste già una Property (Get/Let/Set) con lo stesso nome base
  var existingProperty = mod.Procedures.FirstOrDefault(p => 
      p != proc &&
      p.Kind.StartsWith("Property", StringComparison.OrdinalIgnoreCase) &&
      p.Name.Equals(basePropName, StringComparison.OrdinalIgnoreCase));
  
  if (existingProperty != null && !string.IsNullOrEmpty(existingProperty.ConventionalName))
  {
    // Usa lo stesso ConventionalName della Property esistente
    proc.ConventionalName = existingProperty.ConventionalName;
  }
}
```

## ??? **MODIFICHE AL MODELLO**

### VbProcedure - Nuove Proprietà
```csharp
[JsonIgnore]
public int StartLine { get; set; }      // Riga di inizio procedura

[JsonIgnore] 
public int EndLine { get; set; }        // Riga di fine procedura

public bool ContainsLine(int lineNumber)  // Helper method
{
    return lineNumber >= StartLine && lineNumber <= EndLine;
}
```

### VbModule - Nuovo Metodo
```csharp
public VbProcedure? GetProcedureAtLine(int lineNumber)
{
    return Procedures.FirstOrDefault(p => p.ContainsLine(lineNumber));
}

[JsonIgnore]
public bool IsForm => Kind.Equals("frm", StringComparison.OrdinalIgnoreCase);
```

## ?? **ORDINE DI ESECUZIONE**

### FASE 1: Parsing (`Parser.Core.cs`)
1. Scansiona ogni riga del file
2. Identifica inizio procedure (Function/Sub/Property)
3. Imposta `StartLine = originalLineNumber` 
4. Identifica fine procedure (End Function/Sub/Property)
5. Imposta `EndLine = originalLineNumber`

### FASE 2: Risoluzione (`Parser.Resolve.cs`)
1. Le procedure sono già completamente parsate con StartLine/EndLine
2. Ogni funzione di risoluzione scansiona solo le righe della procedura corrente
3. `GetProcedureAtLine()` funziona correttamente per determinare la procedura

## ?? **SISTEMA DI LOCALIZZAZIONE**

### File di Risorse
- `Resources\Strings.resx` - Inglese (default)
- `Resources\Strings.it.resx` - Italiano
- `Resources\Strings.cs` - Accesso type-safe

### Parametri Comando
```bash
VB6MagicBox.exe --lang=en    # Inglese
VB6MagicBox.exe --lang=it    # Italiano
VB6MagicBox.exe -l it        # Forma breve
```

## ?? **RISULTATI**

### ? **PRIMA delle modifiche**:
```json
"References": [
  {
    "Module": "clsActivePolling",
    "Procedure": "Add",
    "LineNumbers": [218, 232, 236, 243]  // ? SBAGLIATO
  },
  {
    "Module": "clsActivePolling", 
    "Procedure": "Class_Initialize",
    "LineNumbers": [218, 232, 236, 243]  // ? DUPLICATO
  }
]
```

### ? **DOPO le modifiche**:
```json
"References": [
  {
    "Module": "clsActivePolling",
    "Procedure": "Add", 
    "LineNumbers": [218, 232]  // ? Solo righe in questa procedura
  },
  {
    "Module": "clsActivePolling",
    "Procedure": "SomeOtherProcedure",
    "LineNumbers": [236, 243]  // ? Solo righe in quest'altra procedura  
  }
]
```

## 📝 **STATO ATTUALE - DICEMBRE 2024**

### ✅ **PROBLEMI COMPLETAMENTE RISOLTI**:
1. **Riferimenti Duplicati** nei campi delle strutture
2. **IndexOutOfRangeException** durante parsing
3. **Logica Errata** per determinare procedura corrente
4. **Warning API Esterne** per Declare Function/Sub
5. **Accesso Campi Array** non rilevato (`array(i).campo`)
6. **Property Get/Let** con nomi duplicati

### 🔥 **FUNZIONALITÀ PRINCIPALI**:
- ✅ **Parsing VB6** completo (Forms, Modules, Classes)
- ✅ **Analisi References** precisa per ogni simbolo
- ✅ **Refactoring Automatico** con validazione
- ✅ **Backup Progressivo** prima delle modifiche
- ✅ **Encoding VB6** (Windows-1252) supportato
- ✅ **Naming Conventions** moderne applicate

### 📊 **PERFORMANCE**:
- ⚡ **Progress Inline** durante elaborazione
- 🎯 **Controlli Sicurezza** per evitare crash
- 🔒 **Validazione References** prima del refactoring
- 💾 **FileShare.Read** per file aperti in IDE

## 🚀 **PROSSIMI REFACTORING SUGGERITI**

1. **Parser.Core.cs** (1008 linee) - Dividere in:
   - `Parser.Core.Lines.cs` - Gestione continuazioni linee
   - `Parser.Core.Procedures.cs` - Parsing procedure
   - `Parser.Core.Variables.cs` - Parsing variabili
   - `Parser.Core.Statements.cs` - Parsing statement

2. **Parser.Naming.cs** (875 linee) - Valutare se dividere ulteriormente

## 🔧 **NOTE TECNICHE**

- **Target Framework**: .NET 10
- **Compatibilità**: VB6 progetti (.vbp)
- **Encoding**: Supporto per caratteri speciali VB6
- **Performance**: Progress inline durante parsing di progetti grandi

## 📋 **TODO PROSSIME FEATURES**: 
- ✅ **Doppia Esecuzione**: Deve essere idempotente (eseguire due volte, nessun cambiamento)
- 🔧 **Verifica Compilazione**: Il codice refactorizzato deve compilare
- 🧪 **Test Altri VBP**: Provare con diversi progetti VB6 → **PO!!**
- 📝 **Aggiunta Type Hints**: Aggiungere `As NomeTipo` dove manca  
- 🎨 **Fix Righe Bianche**: Normalizzare spaziatura
- 📊 **Ordinamento Dim**: Spostare dichiarazioni a inizio procedure, ordine alfabetico
- 🎯 **Indentazione**: Migliorare formattazione codice > ma c'è già il tool in vb6
- 🗑️ **Rimozione Variabili Non Usate**: Estrarre le variabili non usate / commentarle
- 🪄 **Comando MAGIC**: Creare un comando che esegue tutte le modifiche in automatico