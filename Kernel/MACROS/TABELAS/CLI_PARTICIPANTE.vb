'HASH: BE2C5CA4C9D9ADE8688BF73D30C99F1D
'#Uses "*bsShowMessage"
'CLI_PARTICIPANTE

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim QUANTIDADE As Integer
  Dim QTD As Object
  Set QTD = NewQuery
  QTD.Clear
  QTD.Add("SELECT T.QTDPARTICIPANTES      ")
  QTD.Add("  FROM CLI_TURMA T,            ")
  QTD.Add("       CLI_ATENDIMENTO A       ")
  QTD.Add(" WHERE A.TURMA = T.HANDLE      ")
  QTD.Add("   AND A.HANDLE = :ATENDIMENTO ")
  QTD.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("ATENDIMENTO").AsInteger
  QTD.Active = True
  QUANTIDADE = QTD.FieldByName("QTDPARTICIPANTES").AsInteger
  If QUANTIDADE > 0 Then
    Dim TOTAL As Integer
    Dim VERIFICA As Object
    Set VERIFICA = NewQuery
    VERIFICA.Clear
    VERIFICA.Add("SELECT COUNT(HANDLE) TOTAL FROM CLI_PARTICIPANTE WHERE ATENDIMENTO = :ATENDIMENTO")
    VERIFICA.ParamByName("ATENDIMENTO").Value = CurrentQuery.FieldByName("ATENDIMENTO").AsInteger
    VERIFICA.Active = True
    TOTAL = VERIFICA.FieldByName("TOTAL").AsInteger
    If TOTAL > QUANTIDADE Then
      bsShowMessage("A quantidade máxima de participantes desta turma não pode exceder " + _
             Str(QUANTIDADE) + " participantes!", "E")
      CanContinue = False
      Exit Sub
    End If
    Set VERIFICA = Nothing
  End If
  Set QTD = Nothing
End Sub

