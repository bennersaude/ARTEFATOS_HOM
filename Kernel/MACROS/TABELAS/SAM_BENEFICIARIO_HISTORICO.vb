'HASH: 8A9E478A27B99CD8707C262FFDECF5C0
 
'MACRO: SAM_BENEFICIARIO_HISTORICO

Public Sub TABLE_AfterScroll()
  Dim qSQL As Object
  Set qSQL = NewQuery

  qSQL.Clear
  qSQL.Add ("SELECT C.CONTRATO, C.CONTRATANTE ")
  qSQL.Add ("  FROM SAM_CONTRATO C            ")
  qSQL.Add ("  JOIN SAM_BENEFICIARIO B ON (B.CONTRATO = C.HANDLE) ")
  qSQL.Add (" WHERE B.HANDLE = :HANDLE        ")
  qSQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  qSQL.Active = True

  ROTULOCONTRATO.Text = "Contrato: " + qSQL.FieldByName("CONTRATO").AsString + " - " + qSQL.FieldByName("CONTRATANTE").AsString

  If Not CurrentQuery.FieldByName("BENEFICIARIOORIGEM").IsNull Then
    qSQL.Active = False
    qSQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIOORIGEM").AsInteger
    qSQL.Active = True
    ROTULOCONTRATOORIGEM.Text = "Contrato: " + qSQL.FieldByName("CONTRATO").AsString + " - " + qSQL.FieldByName("CONTRATANTE").AsString
  Else
    ROTULOCONTRATOORIGEM.Text = ""
  End If
  Set QSLQ = Nothing
End Sub
