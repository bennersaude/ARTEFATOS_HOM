'HASH: 16B59383F292A09BF718AA332EABF62C
 
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.State <>1 Then
	bsShowMessage("A tabela não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "5" Then
    bsShowMessage("Processo abortado. A rotina não está Processada!","I")
    Exit Sub
  End If

  Dim sql As Object
  Set sql =NewQuery
  sql.Add("UPDATE ANS_CPM SET SITUACAO = '1', ARQUIVO = NULL WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL
  bsShowMessage("Rotina Cancelada!","I")

  Set sql = Nothing

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <>1 Then
	bsShowMessage("A tabela não pode estar em edição", "I")
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
    bsShowMessage("Processo abortado. A rotina não está mais aberta!","I")
    Exit Sub
  End If

  Dim sx As CSServerExec
  Set sx = NewServerExec

  sx.Description = "CPM - Exportar dados dos Prestadores Médicos"
  sx.DllClassName = "Benner.Saude.ANS.Processos.ProcessoCPM"
  sx.SessionVar("HANDLE_CPM") = CurrentQuery.FieldByName("HANDLE").AsString
  sx.Execute

  Dim sql As Object
  Set sql =NewQuery
  sql.Add("UPDATE ANS_CPM SET SITUACAO = '2' WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL
  bsShowMessage("Processo Enviado para Execução no Servidor!","I")

  Set sql = Nothing
  Set sx = Nothing

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub TABLE_AfterPost()
  Dim sql As Object
  Set sql =NewQuery

  sql.Add("UPDATE ANS_CPM SET USUARIOALTERACAO = :USU, DATAALTERACAO = :DATA WHERE HANDLE = :HANDLE")
  sql.ParamByName("USU").Value =CurrentUser
  sql.ParamByName("DATA").Value =ServerDate
  sql.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL

  Set sql = Nothing

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
 	Select Case CommandID
 		Case "BOTAOPROCESSAR"
 			BOTAOPROCESSAR_OnClick
 		Case "BOTAOCANCELAR"
 			BOTAOCANCELAR_OnClick
	End Select
End Sub
