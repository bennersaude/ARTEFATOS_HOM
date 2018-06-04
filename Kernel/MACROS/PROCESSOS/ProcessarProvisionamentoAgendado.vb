'HASH: BCF5BDD6F4C118C12ED678410B1D221E
'#USES "*CriaTabelaTemporariaSqlServer"
Sub Main
  If InStr(SQLServer, "MSSQL")>0 Then
    CriaTabelaTemporariaSqlServer
  End If

  Dim dllCSharp As CSBusinessComponent
  Set dllCSharp = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.Rotinas.Provisionamento.Agendamento,Benner.Saude.Financeiro.Business")
  dllCSharp.Execute("Run")

  Set dllCSharp = Nothing
End Sub
