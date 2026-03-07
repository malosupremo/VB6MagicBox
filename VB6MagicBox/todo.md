## Idempotenza

## default property 
ma se aggiungessimo i .text o .caption almeno?

### preservare i nomi pubblici delle classi?
enum, const, funct, types?... che noia
mettiamo un flag

## i nomi dei moduli
tipo LOGALARM da dove lo prende? dal PDB?

### missing 
Module,Procedure,Name,ConventionalName,Kind
"clsATH3204","","mbolIsPollingComplete","m_IsPollingComplete","GlobalVariable"
"modGlobal","ByteArrayToIntArray","ByteArrayToIntArray","ByteArrayToIntArray","FunctionReturn"
ma è 
Public Function ByteArrayToIntArray(byteArr() As Byte, Optional ByVal inStart As Integer = -1, Optional ByVal inEnd As Integer = 32767) As Integer()



## LblLbQualcosa
         Begin VB.Label LblLbLatWidth 



--- 

#FATTI

## SHARESTRUCT
doppio sharestruct, PO e CALLER e differenze??

### attributo
Private WithEvents m_ComPort As SaxComm       'Oggetto objComPort di SAX
Attribute objComPort.VB_VarHelpID = -1

### M_
... ma che noia un oggetto privato con eventi diventa m_qualcosa_evento... uffffffffffffffffffffff

### I TAB!!!!!!!!!

# LblOmonime tra etichetta e valore... invece di fare rename con 2, aggiungere Value? se caption ha <numero> allora è un valore, se no è un'etichetta... 

 Begin VB.Label LblGasFlow2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "<135>"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2520
         TabIndex        =   62
         Top             =   420
         Width           =   1080
      End
