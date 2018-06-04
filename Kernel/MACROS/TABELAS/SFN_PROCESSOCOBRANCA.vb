'HASH: 3F209E96E429974DFE815095EEB635E6
'Macro: SFN_PROCESSOCOBRANCA
'#Uses "*bsShowMessage"

Public Sub BOTAOEXCLUIRFATURAS_OnClick()
  Dim SQL As Object
  Dim Fatura As String
  Dim Valor As Currency

  On Error GoTo Fim

  If CurrentQuery.State = 3 Or  CurrentQuery.State = 2 Then
     MsgBox("Salve ou cancele a cobrança antes de excluir a fatura.")
    Exit Sub
  End If

  Fatura = InputBox("Numero da Fatura:")

  If Fatura = "" Then
  	MsgBox("Nenhuma fatura informada.  Exclusão cancelada.")
  	Exit Sub
  End If

  Set SQL = NewQuery

  SQL.Add("SELECT HANDLE, VALOR FROM SFN_FATURA")
  SQL.Add(" WHERE CONTAFINANCEIRA = :CONTAFINANCEIRA")
  'SQL.Add("   AND NUMERO = :FATURA")
  SQL.Add("   AND NUMERO = " + Fatura)
  SQL.Add("   AND HANDLE IN (SELECT FATURA FROM SFN_FATURASPROCESSOCOBRANCA WHERE PROCESSOCOBRANCA = :PROCESSOCOBRANCA)")

  'SQL.ParamByName("FATURA").Value = CInt(Fatura)
  SQL.ParamByName("PROCESSOCOBRANCA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("CONTAFINANCEIRA").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
  SQL.Active = True

  If SQL.EOF Then
Fim:
    MsgBox("Não foi possível excluir a fatura de número " + Fatura)
    Set SQL = Nothing
    Exit Sub
  End If

  Fatura = SQL.FieldByName("HANDLE").AsString
  Valor = SQL.FieldByName("VALOR").AsFloat

  SQL.Active = False

  SQL.Clear
  'SQL.Add("DELETE FROM SFN_FATURASPROCESSOCOBRANCA WHERE FATURA = :FATURA AND PROCESSOCOBRANCA = :PROCESSOCOBRANCA")
  SQL.Add("DELETE FROM SFN_FATURASPROCESSOCOBRANCA")
  SQL.Add("WHERE FATURA = " + Fatura)
  SQL.Add("  And PROCESSOCOBRANCA = :PROCESSOCOBRANCA")

  'SQL.ParamByName("FATURA").Value = CInt(Fatura)
  SQL.ParamByName("PROCESSOCOBRANCA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL

  Set SQL = Nothing

  CurrentQuery.Edit
  If CurrentQuery.FieldByName("VALORCOBRADO").AsFloat - Valor > 0 Then
    CurrentQuery.FieldByName("VALORCOBRADO").Value = CurrentQuery.FieldByName("VALORCOBRADO").AsFloat - Valor
  Else
    CurrentQuery.FieldByName("VALORCOBRADO").Value = 0
  End If
  CurrentQuery.Post


End Sub

Public Sub BOTAOINCLUIRFATURAS_OnClick()
  Dim SQL As Object
  Dim Fatura As String
  Dim Valor As Currency

  On Error GoTo Fim

  If CurrentQuery.State = 3 Or  CurrentQuery.State = 2 Then
     MsgBox("Salve ou cancele a cobrança antes de incluir a fatura.")
    Exit Sub
  End If

  Fatura = InputBox("Numero da Fatura:")

  If Fatura = "" Then
  	MsgBox("Nenhuma fatura informada.  Inclusão cancelada.")
  	Exit Sub
  End If


  Set SQL = NewQuery

  SQL.Add("SELECT HANDLE, VALOR FROM SFN_FATURA")
  SQL.Add(" WHERE CONTAFINANCEIRA = :CONTAFINANCEIRA")
  'SQL.Add("   AND NATUREZA = 'C' AND SITUACAO = 'A' AND NUMERO = :FATURA")
  SQL.Add("   AND NATUREZA = 'C' AND SITUACAO = 'A' AND NUMERO = " + Fatura)
  SQL.Add("   AND HANDLE NOT IN (SELECT FATURA FROM SFN_FATURASPROCESSOCOBRANCA WHERE PROCESSOCOBRANCA = :PROCESSOCOBRANCA)")

  'SQL.ParamByName("FATURA").Value = CInt(Fatura)
  SQL.ParamByName("PROCESSOCOBRANCA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("CONTAFINANCEIRA").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
  SQL.Active = True

  If SQL.EOF Then
Fim:
    MsgBox("Não foi possível incluir a fatura de número " + Fatura)
    Set SQL = Nothing
    Exit Sub
  End If

  Fatura = SQL.FieldByName("HANDLE").AsString
  Valor = SQL.FieldByName("VALOR").AsFloat

  SQL.Active = False

  SQL.Clear
  SQL.Add("INSERT INTO SFN_FATURASPROCESSOCOBRANCA (HANDLE, FATURA, PROCESSOCOBRANCA)")
  'SQL.Add("VALUES (:HANDLE, :FATURA, :PROCESSOCOBRANCA)")
  SQL.Add("VALUES (:HANDLE, " + Fatura + " , :PROCESSOCOBRANCA)")

  SQL.ParamByName("HANDLE").Value = NewHandle("SFN_FATURASPROCESSOCOBRANCA")
  'SQL.ParamByName("FATURA").Value = CInt(Fatura)
  SQL.ParamByName("PROCESSOCOBRANCA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL

  Set SQL = Nothing

  CurrentQuery.Edit
  CurrentQuery.FieldByName("VALORCOBRADO").Value = Valor + CurrentQuery.FieldByName("VALORCOBRADO").AsFloat
  CurrentQuery.Post


End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Add("SELECT C.NOME, A.DATAHORA, B.DESCRICAO")
  SQL.Add("  FROM	SFN_ACOESPROCESSOCOBRANCA A, ")
  SQL.Add("       SFN_ACOESCOBRANCA B, ")
  SQL.Add("       Z_GRUPOUSUARIOS C ")
  SQL.Add(" WHERE PROCESSOCOBRANCA = :HANDLE")
  SQL.Add("   AND B.HANDLE = A.ACOESCOBRANCA")
  SQL.Add("   AND C.HANDLE = A.USUARIO")
  SQL.Add("ORDER BY DATAHORA DESC")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    ROTULO1.Text = "Nenhuma ação de cobrança."
    ROTULO2.Text = ""
  Else
    ROTULO1.Text = "Ultima ação de cobrança:"
    ROTULO2.Text = SQL.FieldByName("DESCRICAO").AsString + "   -   " + SQL.FieldByName("NOME").AsString + "   -   " + Format(SQL.FieldByName("DATAHORA").AsDateTime, "dd/MM/yyyy hh:mm:SS")
  End If

  SQL.Active = False
  Set SQL = Nothing

End Sub


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

Public Sub TABLE_Beforepost(CanContinue As Boolean)
  If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And _
      (CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) Then
    bsShowMessage("A Data Final , se informada, deve ser maior ou igual a inicial", "E")
    CanContinue = False
  Else
    CanContinue = True
  End If

End Sub


