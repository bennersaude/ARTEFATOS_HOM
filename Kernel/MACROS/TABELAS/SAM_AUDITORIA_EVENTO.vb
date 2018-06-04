'HASH: 568DF14FD83E014855313D5FCE0D5041
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Exclusão de exceção permitida apenas pela interface de Autorização", "E")
  End If
End Sub



Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Alteração de exceção permitida apenas pela interface de Autorização", "E")
  End If

  Dim SQL As String
  SQL = "SELECT D.HANDLE " + _
        "  FROM SAM_AUDITORIA A, " + _
        "       SAM_AUTORIZ B, " + _
        "       SAM_AUTORIZ_EVENTOSOLICIT C, " + _
        "       SAM_AUTORIZ_EVENTOGERADO D " + _
        " WHERE B.HANDLE = A.AUTORIZACAO " + _
        "   AND C.AUTORIZACAO = B.HANDLE " + _
        "   AND D.EVENTOSOLICITADO = C.HANDLE " + _
        "   AND A.HANDLE = " + Str(RecordHandleOfTable("SAM_AUDITORIA")) + _
        "   AND D.SITUACAO IN ('L','A') "

  If WebMode Then
  	 EVENTOGERADO.WebLocalWhere = "A.HANDLE IN (" + SQL + ")"
  ElseIf VisibleMode Then
 	 EVENTOGERADO.LocalWhere = "SAM_AUTORIZ_EVENTOGERADO.HANDLE IN (" + SQL + ")"
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Inclusão de exceção permitida apenas pela interface de Autorização", "E")
  End If

  Dim SQL As String
  SQL = "SELECT D.HANDLE " + _
        "  FROM SAM_AUDITORIA A, " + _
        "       SAM_AUTORIZ B, " + _
        "       SAM_AUTORIZ_EVENTOSOLICIT C, " + _
        "       SAM_AUTORIZ_EVENTOGERADO D " + _
        " WHERE B.HANDLE = A.AUTORIZACAO " + _
        "   AND C.AUTORIZACAO = B.HANDLE " + _
        "   AND D.EVENTOSOLICITADO = C.HANDLE " + _
        "   AND A.HANDLE = " + Str(RecordHandleOfTable("SAM_AUDITORIA")) + _
        "   AND D.SITUACAO IN ('L','A') "

  If WebMode Then
  	 EVENTOGERADO.WebLocalWhere = "A.HANDLE IN (" + SQL + ")"
  ElseIf VisibleMode Then
 	 EVENTOGERADO.LocalWhere = "SAM_AUTORIZ_EVENTOGERADO.HANDLE IN (" + SQL + ")"
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString = "G" And _
                              CurrentQuery.FieldByName("MOTIVOGLOSA").IsNull Then
    bsShowMessage("Ação exige o preenchimento do campo motivo de glosa.", "E")
    CanContinue = False
  End If

  If CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString = "O" And _
                              CurrentQuery.FieldByName("OBSERVACOES").IsNull Then
    bsShowMessage("Ação exige o preenchimento do campo OBSERVAÇÃO.", "E")
    CanContinue = False
  End If

End Sub

