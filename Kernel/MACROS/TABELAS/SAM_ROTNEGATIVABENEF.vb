'HASH: 071B7FF3B50E6793FE2D9BFA88BD82AA
'#Uses "*bsShowMessage"

Public Sub AUTORIZACAO_OnPopup(ShowPopup As Boolean)
Dim ProcuraDLL As Object
  Dim handlexx As Long
  ShowPopup = False
  Dim vPos As Integer

  Dim SQL As Object
  Set SQL = NewQuery

  On Error GoTo prox
    SQL.Add("SELECT HANDLE FROM SAM_AUTORIZ WHERE AUTORIZACAO=:AUTORIZ")
    SQL.ParamByName("AUTORIZ").AsString = AUTORIZACAO.Text
    SQL.Active = True
    If Not SQL.EOF Then
      CurrentQuery.FieldByName("AUTORIZACAO").AsInteger = SQL.FieldByName("HANDLE").AsInteger
      Set SQL = Nothing
      Exit Sub
    End If
prox:
  Set SQL = Nothing
  Set ProcuraDLL = CreateBennerObject("Procura.Procurar")

  If (IsNumeric(AUTORIZACAO.Text)) Then
    vPos = 2 'realiza a busca por autorizacao
  Else
    vPos = 4 'realiza a busca por beneficiario
  End If

  handlexx = ProcuraDLL.Exec(CurrentSystem, "SAM_AUTORIZ|SAM_BENEFICIARIO[SAM_BENEFICIARIO.HANDLE=SAM_AUTORIZ.BENEFICIARIO]", "SAM_AUTORIZ.DATAAUTORIZACAO|SAM_AUTORIZ.AUTORIZACAO|SAM_AUTORIZ.RADIOSOLICITACAO|SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_AUTORIZ.SENHAGUIASOLICITACAO", vPos, "Data da Autorização|Autorização|Solicitação|Nome|Matrícula Funcional|Senha da solicitação", "SAM_AUTORIZ.HANDLE IN (SELECT B.AUTORIZACAO FROM SAM_AUTORIZ_EVENTOGERADO B WHERE B.SITUACAO = 'N') ", "Procura Autorização", False, AUTORIZACAO.Text)


  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("AUTORIZACAO").Value = handlexx
  End If
  Set ProcuraDLL = Nothing
End Sub

Public Sub BOTAOAUTORIZACAO_OnClick()

  If CurrentQuery.State <> 1 Then
    bsshowmessage("Registro está em edição", "I")
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("CA043.Autorizacao")
  Interface.Executar(CurrentSystem, 0, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, 0)
  Set Interface = Nothing
End Sub
