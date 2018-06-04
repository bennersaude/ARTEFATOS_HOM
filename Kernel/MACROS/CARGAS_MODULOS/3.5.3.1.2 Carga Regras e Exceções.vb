'HASH: 34287A0E60C07E0BF3A678AFC18F32B3
Public Sub BOTAOGERAREVENTOS_OnClick()
Dim interface As Object
Dim SQL As Object
Set SQL =NewQuery
SQL.Add("SELECT DATAFINAL FROM SAM_PRESTADOR_PROC WHERE HANDLE = :HANDLE")
SQL.ParamByName("HANDLE").Value =RecordHandleOfTable("SAM_PRESTADOR_PROC")
SQL.Active =True
If SQL.FieldByName("DATAFINAL").IsNull Then
Set interface =CreateBennerObject("SAMPROCPRESTADOR.PROCESSOPRESTADOR")
If InTransaction Then
Rollback
End If
interface.GerarEventos(CurrentSystem,RecordHandleOfTable("SAM_PRESTADOR"),RecordHandleOfTable("SAM_PRESTADOR_PROC_CREDEN"),CurrentUser)
Set interface =Nothing
RefreshNodesWithTable "SAM_PRESTADOR_PROC_REGEXC"
Else
MsgBox "Processo finalizado operação Não Permitida"
End If
SQL.Active =False
Set SQL =Nothing
End Sub
