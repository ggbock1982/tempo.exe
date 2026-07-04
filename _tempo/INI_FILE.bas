Attribute VB_Name = "INI_FILE"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long
    
    
Public Function LerINI(ByVal Secao As String, ByVal Chave As String, ByVal Arquivo As String) As String
  Dim Retorno As String
  Dim ValorINI As Long
  Retorno = Space$(255)
  ValorINI = GetPrivateProfileString(Secao, Chave, "", Retorno, 255, Arquivo)
  LerINI = Left$(Retorno, ValorINI)
End Function

Public Sub GravarINI(ByVal Secao As String, ByVal Chave As String, ByVal Valor As String, ByVal Arquivo As String)
  WritePrivateProfileString Secao, Chave, Valor, Arquivo
End Sub
