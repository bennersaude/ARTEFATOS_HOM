'HASH: 7B82D4DB87606210188BCC6F8F41D4C7
'#USES "*CriaTabelaTemporariaSqlServer"
Sub Main
  Dim dll As Object
 If InStr(SQLServer, "MSSQL")>0 Then
        CriaTabelaTemporariaSqlServer
 end if
  Set dll=CreateBennerObject("sampeg.processar")
  dll.ProcessarAgendamento(CurrentSystem,CurrentUser,1)
  Set dll=Nothing
End sub
