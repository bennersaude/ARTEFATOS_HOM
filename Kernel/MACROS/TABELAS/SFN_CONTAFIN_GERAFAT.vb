'HASH: 292D6C6CB6973DE40218F5A3FE9C8154
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Q1 As Object
  Dim Q2 As Object
  Set Q1 = NewQuery
  Set Q2 = NewQuery

  Q1.Add("SELECT SUM(PERCENTUAL) TOTAL")
  Q1.Add("  FROM SFN_TIPOFATURA_CLASSEGER")
  Q1.Add(" WHERE TIPOFATURA = :TIPOFATURA")
  Q1.Add("HAVING SUM(PERCENTUAL) <> 100")
  Q1.ParamByName("TIPOFATURA").Value = CurrentQuery.FieldByName("TIPOFATURA").AsInteger
  Q1.Active = True

  If Not Q1.EOF Then
    bsShowMessage("Verificar tipo de fatura." + Chr(13) + "A soma dos percentuais da classe gerencial deve ser 100%", "E")
    CanContinue = False
    Set Q1 = Nothing
    Set Q2 = Nothing
    Exit Sub
  End If

  Q2.Add("SELECT SUM(PERCENTUAL) TOTAL")
  Q2.Add("  FROM SFN_TIPOFATURA_CLASSEGER_CC")
  Q2.Add(" WHERE TIPOFATURACLASSEGER IN (SELECT HANDLE")
  Q2.Add("                                 FROM SFN_TIPOFATURA_CLASSEGER")
  Q2.Add("                                WHERE TIPOFATURA = :TIPOFATURA)")
  Q2.Add(" GROUP BY TIPOFATURACLASSEGER")
  Q2.Add("HAVING SUM(PERCENTUAL) <> 100")

  Q2.ParamByName("TIPOFATURA").Value = CurrentQuery.FieldByName("TIPOFATURA").AsInteger
  Q2.Active = True

  If Not Q2.EOF Then
    bsShowMessage("Verificar tipo de fatura." + Chr(13) + "A soma dos percentuais do centro de custo deve ser 100%", "E")
    CanContinue = False
    Set Q1 = Nothing
    Set Q2 = Nothing
    Exit Sub
  End If

  Set Q1 = Nothing
  Set Q2 = Nothing

  If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And _
      (CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) Then
    bsShowMessage("A Data Final , se informada, deve ser maior ou igual a inicial", "E")
    CanContinue = False
  Else
    CanContinue = True
  End If


End Sub

