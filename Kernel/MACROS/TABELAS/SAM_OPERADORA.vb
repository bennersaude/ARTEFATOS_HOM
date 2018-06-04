'HASH: AFB77F76BA3B27F627CA26A53D5AB110
'#Uses "*bsShowMessage"

Public Sub ADMINISTRADORA_OnPopup(ShowPopup As Boolean)
  ADMINISTRADORA.LocalWhere = "TABTIPO = 1"

End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABTIPO_OnChange()
  Dim SQL As Object
  Set SQL = NewQuery

  If TABTIPO.PageIndex = 1 Then
    SQL.Clear
    SQL.Add("UPDATE SAM_OPERADORA SET ADMINISTRADORA = NULL WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT HANDLE")
    SQL.Add("FROM SAM_OPERADORA")
    SQL.Add("WHERE NUMEROREGISTRO = :NRREGISTRO")
    SQL.ParamByName("NRREGISTRO").AsInteger = CurrentQuery.FieldByName("NUMEROREGISTRO").AsString
    SQL.Active = True

    If Not SQL.EOF Then
      BsShowMessage("Já existe este Número de Registro cadastrado!", "E")
      CanContinue = False
	  Exit Sub
    End If

    Set SQL = Nothing
  End If

  If Not IsValidCGC(CurrentQuery.FieldByName("CNPJ").AsString)Then
    BsShowMessage("CNPJ Inválido", "E")
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABTIPO").AsString = "2" And Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
      BsShowMessage("A data final não pode ser menor do que a data inicial", "E")
      CanContinue = False
      Exit Sub
    End If
  End If


  Dim vRegistroAns As String

  If (CurrentQuery.FieldByName("REGISTROANS").IsNull) Then
    vRegistroAns = CurrentQuery.FieldByName("NUMEROREGISTRO").AsString+CurrentQuery.FieldByName("DIGITOVERIFICADOR").AsString

    While (Len(vRegistroAns) < 6)
      vRegistroAns = "0"+vRegistroAns
    Wend

    CurrentQuery.FieldByName("REGISTROANS").AsString = vRegistroAns
  End If


End Sub
