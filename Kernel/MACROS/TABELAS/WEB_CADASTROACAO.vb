'HASH: FA44483EC9B608E3EF2DCC3ACFF6554B
 

Public Sub SITUACAORH_OnPopup(ShowPopup As Boolean)
  SITUACAORH.LocalWhere = "TABTIPOSITUACAO = '2'"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT COUNT(1) QTD")
  SQL.Add("  FROM WEB_CADASTROACAO     A")
  SQL.Add(" WHERE A.IDENTIFICADORVISAO = :IDENTIFICADOR")
  SQL.Add("   AND A.WEBCADASTROPROCESSO <> :WEBCADASTROPROCESSO")
  SQL.ParamByName("IDENTIFICADOR").AsString = CurrentQuery.FieldByName("IDENTIFICADORVISAO").AsString
  SQL.ParamByName("WEBCADASTROPROCESSO").AsInteger = CurrentQuery.FieldByName("WEBCADASTROPROCESSO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("QTD").AsInteger > 0 Then
    CanContinue = False
    MsgBox "A visão associada está associada a outro processo web. Operação não permitida."
  End If



End Sub
