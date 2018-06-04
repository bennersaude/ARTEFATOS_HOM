'HASH: DD77F0B53FFC435DBB79E56E9227A737
 
'#Uses "*bsShowMessage"
Option Explicit


Public Sub ITEM_OnPopup(ShowPopup As Boolean)
  'ITEM.LocalWhere = "LICITACAO = " + Str(RecordHandleOfTable("SAM_LICITACAO")) 'SMS 167330
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery
  sql.Add("SELECT COUNT(1) QTD ")
  sql.Add("  FROM SAM_LICITACAO_FORNECEDOR_ITEM")
  sql.Add(" WHERE LICITACAOFORNECEDOR = :LICITACAO")
  sql.Add("   AND ITEM = :ITEM")
  sql.Add("   AND HANDLE <> :HANDLE")
  sql.ParamByName("LICITACAO").AsInteger = CurrentQuery.FieldByName("LICITACAOFORNECEDOR").AsInteger
  sql.ParamByName("ITEM").AsInteger = CurrentQuery.FieldByName("ITEM").AsInteger
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True
  If sql.FieldByName("QTD").AsInteger > 0 Then
    BsShowMessage("Item existe na cotação. Não é possível inserí-lo novamente", "E")
    CanContinue = False
  End If

End Sub
