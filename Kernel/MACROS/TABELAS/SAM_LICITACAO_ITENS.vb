'HASH: 5C78F710F54870D804A35216F65297EA

'#Uses "*bsShowMessage"
'#Uses "*ProcuraEventoMatMed"

Public Function VerificarCotacaoIniciada(pLicitacao As Long) As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT PUBLICADAWEB FROM SAM_LICITACAO WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = pLicitacao
  SQL.Active = True

  If SQL.FieldByName("PUBLICADAWEB").AsString = "S" Then
    VerificarCotacaoIniciada = True
  End If
  Set SQL=Nothing
End Function


Public Sub MATMED_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEventoMatMed(True, MATMED.Text)

  If vHandle <> 0 Then
	CurrentQuery.Edit
	CurrentQuery.FieldByName("MATMED").Value = vHandle
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If VerificarCotacaoIniciada(RecordHandleOfTable("SAM_LICITACAO")) Then
    CanContinue = False
    BsShowMessage("Cotação já iniciada. Não é possível excluir materiais/medicamentos.", "E")

  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If VerificarCotacaoIniciada(RecordHandleOfTable("SAM_LICITACAO")) Then
    CanContinue = False
    BsShowMessage("Cotação já iniciada. Não é possível editar materiais/medicamentos.", "E")
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If VerificarCotacaoIniciada(RecordHandleOfTable("SAM_LICITACAO")) Then
    CanContinue = False
    BsShowMessage("Cotação já iniciada. Não é possível acrescentar materiais/medicamentos.", "E")
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
'  If CurrentQuery.FieldByName("PRECOMAXIMO").AsFloat <= 0 Then
'    CanContinue = False
'    If VisibleMode Then
'      MsgBox "Preço máximo deve ser maior que zero."
'    End If
'  End If
'TJDF- O cliente solicita que nao necessite informar preço máximo

  If CurrentQuery.FieldByName("QUANTIDADE").AsInteger <= 0 Then
    CanContinue = False
    BsShowMessage("Quantidade solicitada deve ser maior que zero.", "E")
  End If



End Sub
