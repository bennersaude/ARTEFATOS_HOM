'HASH: 28882EBC806200717E84B3848FFEB833
'TABELA: AT_ATEND_REALIZADOMED

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub



Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  If CurrentQuery.FieldByName("EVENTO").IsNull Then
    bsShowMessage("É necessário o evento!", "I")
    Exit Sub
  End If

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String

  Set interface = CreateBennerObject("Procura.Procurar")

  vTabela = "SAM_GRAU"
  vColunas = "SAM_GRAU.DESCRICAO|SAM_GRAU.XTHM|SAM_GRAU.VERIFICAGRAUSVALIDOS"
  vCriterio = "SAM_GRAU.VERIFICAGRAUSVALIDOS = 'N' " + _
              "OR (EXISTS (SELECT GE.HANDLE FROM SAM_TGE_GRAU GE WHERE GE.EVENTO=" + _
              CurrentQuery.FieldByName("EVENTO").AsString + _
              " AND GE.GRAU=SAM_GRAU.HANDLE))"
  vCampos = "Descrição|XTHM|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Graus válidos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("QUANTIDADE").AsInteger <= 0 Then
    bsShowMessage("A quantidade deve ser no mínimo 1!", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim myQuery As Object

  Set myQuery = NewQuery
  myQuery.Active = False
  myQuery.Add("SELECT DATAATENDIMENTO,HORAATENDIMENTO FROM AT_ATEND WHERE HANDLE=:pHANDLE")
  myQuery.ParamByName("pHANDLE").Value = CurrentQuery.FieldByName("Atendimento").Value
  myQuery.Active = True
  If(myQuery.FieldByName("DATAATENDIMENTO").IsNull)Or(myQuery.FieldByName("HORAATENDIMENTO").IsNull)Then
  bsShowMessage("Atendimento não foi realizado.Os campos Data e Hora do Atendimento deverão ser preenchidos!", "E")
  CanContinue = False
End If
Set myQuery = Nothing

End Sub

