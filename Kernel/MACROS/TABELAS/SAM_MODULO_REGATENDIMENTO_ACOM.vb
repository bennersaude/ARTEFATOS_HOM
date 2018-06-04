'HASH: E73064710771E7D7FF97EF9B84A7FF41
'MACRO: SAM_MODULO_REGATENDIMENTO_ACOM
'#Uses "*bsShowMessage"

'SMS 55871 - Matheus - 08/08/2006
Public Sub ACOMODACAO_OnPopup(ShowPopup As Boolean)
  ACOMODACAO.LocalWhere = "HANDLE NOT IN (SELECT ACOMODACAO FROM SAM_MODULO_REGATENDIMENTO_ACOM "+ _
  "WHERE MODULOREGATENDIMENTO = "+CurrentQuery.FieldByName("MODULOREGATENDIMENTO").AsString+")"
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim QACOMODACAO As Object
  Set QACOMODACAO = NewQuery

  QACOMODACAO.Active = False
  QACOMODACAO.Clear
  QACOMODACAO.Add("SELECT COUNT(*) QTDE                     ")
  QACOMODACAO.Add("  FROM SAM_MODULO_REGATENDIMENTO_ACOM    ")
  QACOMODACAO.Add(" WHERE MODULOREGATENDIMENTO = :MODULOREG ")
  QACOMODACAO.ParamByName("MODULOREG").AsInteger = CurrentQuery.FieldByName("MODULOREGATENDIMENTO").AsInteger
  QACOMODACAO.Active = True

  If (QACOMODACAO.FieldByName("QTDE").AsInteger > 1) Then
    If (CurrentQuery.FieldByName("PADRAO").AsString = "S") Then
      bsShowMessage("Não é permitido excluir a acomodação padrão!","E")
      CanContinue = False
      Set QACOMODACAO = Nothing
      Exit Sub
    End If
  End If

  Set QACOMODACAO = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim QACOMODACAO As Object
  Set QACOMODACAO = NewQuery

  QACOMODACAO.Active = False
  QACOMODACAO.Clear
  QACOMODACAO.Add("SELECT HANDLE                            ")
  QACOMODACAO.Add("  FROM SAM_MODULO_REGATENDIMENTO_ACOM    ")
  QACOMODACAO.Add(" WHERE PADRAO = 'S'                      ")
  QACOMODACAO.Add("   AND MODULOREGATENDIMENTO = :MODULOREG ")
  QACOMODACAO.Add("   AND HANDLE <> :HANDLE                 ")
  QACOMODACAO.ParamByName("MODULOREG").AsInteger = CurrentQuery.FieldByName("MODULOREGATENDIMENTO").AsInteger
  QACOMODACAO.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QACOMODACAO.Active = True

  If (CurrentQuery.FieldByName("PADRAO").AsString = "N") And (QACOMODACAO.EOF) Then
    bsShowMessage("Deve existir uma acomodação padrão para este regime de atendimento!","E")
    CanContinue = False
    Set QACOMODACAO = Nothing
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("PADRAO").AsString = "S") And (Not QACOMODACAO.EOF) Then

    If (bsShowMessage("Já existe uma acomodação padrão para este regime de atendimento!" + Chr(13) + "Deseja que esta acomodação passe a ser a padrão?", "Q") = vbYes) Then

      Dim QMUDAPADRAO As Object
      Set QMUDAPADRAO = NewQuery
      QMUDAPADRAO.Active = False
      QMUDAPADRAO.Clear
      QMUDAPADRAO.Add("UPDATE SAM_MODULO_REGATENDIMENTO_ACOM ")
      QMUDAPADRAO.Add("   SET PADRAO = 'N'                   ")
      QMUDAPADRAO.Add(" WHERE HANDLE = :HANDLE               ")
      QMUDAPADRAO.ParamByName("HANDLE").AsInteger = QACOMODACAO.FieldByName("HANDLE").AsInteger
      QMUDAPADRAO.ExecSQL

      Set QMUDAPADRAO = Nothing
    End If
  End If

  Set QACOMODACAO = Nothing
End Sub
'Fim SMS 55871
