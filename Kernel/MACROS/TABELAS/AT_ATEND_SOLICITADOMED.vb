'HASH: DEC3E60AE6CF269B5F5DE4930B1F92BA
'Tabela: AT_ATEND_SOLICITADOMED

'#Uses "*ProcuraEvento"

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
    MsgBox("É necessário o evento!")
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

