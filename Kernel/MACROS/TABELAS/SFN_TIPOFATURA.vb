'HASH: F5F76D7AC1053179D5149A5E70E372F9
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	'------------------- SMS 90104 - Paulo Melo - 17/12/2007 - INICIO  -- Fazer o tratamento para nao dar erro de unique do banco, mas sim mostra uma mensagem.
 Dim CODIGO As Object
 Set CODIGO = NewQuery

 CODIGO.Add("  SELECT HANDLE            ")
 CODIGO.Add("    FROM SFN_TIPOFATURA    ")
 CODIGO.Add("   WHERE CODIGO = :CODIGO  ")
 CODIGO.Add("     AND HANDLE <> :HANDLE ")
 CODIGO.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
 CODIGO.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 CODIGO.Active = True

 If Not CODIGO.EOF Then
 	bsShowMessage("Já existe uma fatura com esse código", "E")
 	CanContinue = False
 End If

 Set CODIGO = Nothing
'------------------- SMS 90104 - Paulo Melo - 17/12/2007 - FIM
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  'SMS 43004
  If (VisibleMode And NodeInternalCode = 43004) Or (WebMode And WebMenuCode = "T5431") Then
    If WebMode Then
      TIPOFATURAMENTO.WebLocalWhere = " (A.CODIGO = 130 OR A.CODIGO = 310 OR A.CODIGO = 410)"
    Else
      TIPOFATURAMENTO.LocalWhere    = " (CODIGO = 130 OR CODIGO = 310 OR CODIGO = 410)"
    End If
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  'SMS 43004
  If (VisibleMode And NodeInternalCode = 43004) Or (WebMode And WebMenuCode = "T5431") Then
    If WebMode Then
      TIPOFATURAMENTO.WebLocalWhere = " (A.CODIGO = 130 OR A.CODIGO = 310 OR A.CODIGO = 410)"
    Else
      TIPOFATURAMENTO.LocalWhere    = " (CODIGO = 130 OR CODIGO = 310 OR CODIGO = 410)"
    End If
  End If
End Sub
