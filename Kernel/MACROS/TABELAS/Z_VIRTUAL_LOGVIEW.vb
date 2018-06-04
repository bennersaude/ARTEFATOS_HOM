'HASH: EEA7BEA8F3B3F8FD90BE365D54052650
Option Explicit 
 
Public Sub TABLE_OnVirtualValue(ByVal MacroName As String, ResultValue As String) 
 
  Select Case MacroName 
    Case Is = "USERFILTER" 
      Rem ************************************************************ 
      Rem Retornar um SQLEspecial para filtro dos registros de Z_LOG * 
      Rem Exemplo: "APELIDO LIKE 'RUNNER%' AND HANDLE <> 1"          * 
      Rem ************************************************************ 
 
    Case Is = "USEROPERATOR" 
      Rem ************************************************************ 
      Rem Retornar um operador (AND, OR) que deve ser utilizado      * 
      Rem no SQLEspecial retornado por "USERFILTER"                  * 
      Rem ************************************************************ 
 
  End Select 
 
End Sub 
