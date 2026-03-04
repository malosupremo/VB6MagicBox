### attributo
Private WithEvents m_ComPort As SaxComm       'Oggetto objComPort di SAX
Attribute objComPort.VB_VarHelpID = -1

... ma che noia un oggetto privato con eventi diventa m_qualcosa_evento... uffffffffffffffffffffff

### preservare i nomi pubblici delle classi?
enum, const, funct, types?... che noia


### missing 
Module,Procedure,Name,ConventionalName,Kind
"clsATH3204","","mbolIsPollingComplete","m_IsPollingComplete","GlobalVariable"
"modGlobal","ByteArrayToIntArray","ByteArrayToIntArray","ByteArrayToIntArray","FunctionReturn"
ma è 
Public Function ByteArrayToIntArray(byteArr() As Byte, Optional ByVal inStart As Integer = -1, Optional ByVal inEnd As Integer = 32767) As Integer()