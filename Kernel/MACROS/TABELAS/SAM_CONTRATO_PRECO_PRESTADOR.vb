'HASH: 34012A5628D442D2498823C7E3321BAD
'#Uses "*bsShowMessage"

Option Explicit

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  '#Uses "*ProcuraPrestador"
  '  If Len(PRESTADOR.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", PRESTADOR.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  '  End If
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  sql.Clear
  'Verificação para ver se vai sobrar algum prestador
  sql.Add("SELECT COUNT(HANDLE) QTDE FROM SAM_CONTRATO_PRECO_PRESTADOR ")
  sql.Add("WHERE CONTRATOPRECO = " + CurrentQuery.FieldByName("CONTRATOPRECO").AsString)
  sql.Active = True
  If sql.FieldByName("QTDE").AsInteger = 1 Then 'É o último prestador da tabela
    Dim SQL2 As Object
    Set SQL2 = NewQuery

    'MsgBox("Último da tabela")
    'Busca a que tabela e que contrato o prestador pertence
    sql.Clear
    sql.Add ("SELECT TABELAPRECO, CONTRATO, ESTRUTURAINICIAL, ESTRUTURAFINAL      ")
    sql.Add ("  FROM SAM_CONTRATO_PRECO WHERE HANDLE = " + CurrentQuery.FieldByName("CONTRATOPRECO").AsString)
    sql.Active = True

    SQL2.Clear
    SQL2.Add ("SELECT TABELAPRECO, CONTRATO , ESTRUTURAINICIAL, ESTRUTURAFINAL ")
    SQL2.Add ("  FROM SAM_CONTRATO_PRECO                                       ")
    SQL2.Add (" WHERE ")
    'SQL2.ADD (" TABELAPRECO = " + sql.FieldByName("TABELAPRECO").AsString)
    'SQL2.Add ("   And CONTRATO    = " + sql.FieldByName("CONTRATO").AsString)
    SQL2.Add ("      CONTRATO    = " + sql.FieldByName("CONTRATO").AsString)
    SQL2.Add ("   And (Replace(ESTRUTURAINICIAL, '.','') BETWEEN REPLACE(:ESTRUTURAINICIAL,'.','')")
    SQL2.Add ("                                              And REPLACE(:ESTRUTURAFINAL  ,'.','')")
    SQL2.Add ("    Or  Replace(ESTRUTURAFINAL  , '.','') BETWEEN REPLACE(:ESTRUTURAINICIAL,'.','')")
    SQL2.Add ("                                              And REPLACE(:ESTRUTURAFINAL  ,'.','')")
    SQL2.Add (")")
    SQL2.Add ("   And HANDLE Not In (Select CONTRATOPRECO FROM SAM_CONTRATO_PRECO_PRESTADOR) ")
    SQL2.ParamByName ("ESTRUTURAINICIAL").Value = sql.FieldByName("ESTRUTURAINICIAL").AsString
    SQL2.ParamByName ("ESTRUTURAFINAL").Value = sql.FieldByName("ESTRUTURAFINAL").AsString
    'MsgBox(SQL2.Text)
    SQL2.Active = True

    If SQL2.FieldByName("CONTRATO").IsNull Then
      bsShowMessage("Esta tabela de cobrança, será utilizada para todos os prestadores nesta faixa de evento", "I")
    Else
      bsShowMessage("Já possui uma tabela para todos os prestadores, pos isto nao pode ser excluido", "E")
      CanContinue = False
      Set SQL2 = Nothing
    End If
  End If
End Sub



Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sqlVerificaExistencia As Object
  Set sqlVerificaExistencia = NewQuery

  sqlVerificaExistencia.Clear
  sqlVerificaExistencia.Add("SELECT CONTRATOPRECO                                            ")
  sqlVerificaExistencia.Add("  FROM SAM_CONTRATO_PRECO_PRESTADOR CPP                         ")
  sqlVerificaExistencia.Add("  JOIN SAM_CONTRATO_PRECO CP On CP.Handle = CPP.CONTRATOPRECO   ")
  sqlVerificaExistencia.Add(" WHERE CP.CONTRATO = (SELECT CONTRATO                           ")
  sqlVerificaExistencia.Add("                        FROM SAM_CONTRATO_PRECO                 ")
  sqlVerificaExistencia.Add("                       WHERE HANDLE = :CONTRATOPRECO)           ")
  sqlVerificaExistencia.Add("   AND CPP.Prestador = :PRESTADOR                               ")
  sqlVerificaExistencia.Add("   AND CPP.Handle <> :HANDLE                                    ")
  sqlVerificaExistencia.ParamByName("CONTRATOPRECO").Value = CurrentQuery.FieldByName("CONTRATOPRECO").AsString
  sqlVerificaExistencia.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
  sqlVerificaExistencia.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  sqlVerificaExistencia.Active = True

  If Not sqlVerificaExistencia.EOF Then
    If sqlVerificaExistencia.FieldByName("CONTRATOPRECO").AsString = CurrentQuery.FieldByName("CONTRATOPRECO").AsString Then
      bsShowMessage("Este Prestador já está relacionado nesta tabela de preço", "E")
    Else
      bsShowMessage("Este Prestador já está relacionada em outra tabela de preço", "E")
    End If
    CanContinue = False
  End If
  Set sqlVerificaExistencia = Nothing
End Sub
