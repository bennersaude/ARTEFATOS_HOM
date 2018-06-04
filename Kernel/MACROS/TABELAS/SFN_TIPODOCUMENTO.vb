'HASH: FBC07B083263B783EF3555627EA66C75
'Macro: SFN_TIPODOCUMENTO
'#Uses "*bsShowMessage"


Public Sub BOTAOMODELODOCUMENTO_OnClick()

  Dim Obj As Object
  Set Obj = CreateBennerObject("SFNDOCUMENTO.ROTINAS")
  Obj.CadModeloTipo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Obj.Finalizar(CurrentSystem)
  Set Obj = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  UserVar("ROTINAARQUIVO")  = "TESTE RELATORIO"
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim sqlOrdem As Object
  Set sqlOrdem = NewQuery

  sqlOrdem.Clear
  sqlOrdem.Add("SELECT COUNT(*) QTDE FROM SFN_TIPODOCUMENTO_BAIXA WHERE TIPODOCUMENTO =:HTIPODOC")
  sqlOrdem.ParamByName("HTIPODOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sqlOrdem.Active = True

  If sqlOrdem.FieldByName("QTDE").AsInteger > 0 Then
    If visibilemode Then
    	Beep
    End If
    bsShowMessage("Operação Cancelada! - Existe(m) tipo(s) de fatura cadastrada(s) na Ordem para baixa de faturas!", "E")
    sqlOrdem.Active = False
    Set sqlOrdem = Nothing
    CanContinue = False
    Exit Sub
  End If

  Dim sqlBanco As Object
  Set sqlBanco = NewQuery

  sqlBanco.Clear
  sqlBanco.Add("SELECT COUNT(*) QTDE FROM SFN_TIPODOCUMENTO_BANCO WHERE TIPODOCUMENTO =:HTIPODOC")
  sqlBanco.ParamByName("HTIPODOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sqlBanco.Active = True

  If sqlBanco.FieldByName("QTDE").AsInteger > 0 Then
	If VisibleMode Then
    	Beep
    End If
    bsShowMessage("Operação Cancelada! - Existe(m) Banco(s) ligado(s) a este tipo de documento!", "E")
    sqlBanco.Active = False
    Set sqlBanco = Nothing
    CanContinue = False
    Exit Sub
  End If

  Dim sqlFaturamento As Object
  Set sqlFaturamento = NewQuery

  sqlFaturamento.Clear
  sqlFaturamento.Add("SELECT COUNT(*) QTDE FROM SFN_TIPODOCUMENTO_TIPOFATURAME WHERE TIPODOCUMENTO =:HTIPODOC")
  sqlFaturamento.ParamByName("HTIPODOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sqlFaturamento.Active = True

  If sqlFaturamento.FieldByName("QTDE").AsInteger > 0 Then
    If VisibleMode Then
    	Beep
    End If
    bsShowMessage("Operação Cancelada! - Existe(m) Tipo(s) de Faturamento ligado(s) a este tipo de documento!", "E")
    sqlFaturamento.Active = False
    Set sqlFaturamento = Nothing
    CanContinue = False
    Exit Sub
  End If

  sqlOrdem.Active = False
  sqlBanco.Active = False
  sqlFaturamento.Active = False

  Set sqlOrdem = Nothing
  Set sqlBanco = Nothing
  Set sqlFaturamento = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("CONTAFINTABGERACAO").AsInteger = 3 And CurrentQuery.FieldByName("TABTIPO").AsInteger<3 Then
    bsShowMessage("Conta financeira com Título não pode ter tipo de documento com movimento bancário", "E")
    CanContinue = False
  End If

  If CurrentQuery.FieldByName("CONTAFINTABGERACAO").AsInteger = 1 And CurrentQuery.FieldByName("TABTIPO").AsInteger>= 3 Then
    bsShowMessage("Conta financeira com conta corrente não pode ter tipo de documento diferente de conta correte ou DOC", "E")
    CanContinue = False
  End If

  'SMS 21072 - Leonam - 28/11/2003
  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 3 And CurrentQuery.FieldByName("CARTEIRA").IsNull Then
    bsShowMessage("Campo CARTEIRA deve ser preenchido quando conta financeira do tipo Boleto.", "E")
    CanContinue = False
  End If

  'SMS 19516 - Kristian - 06/04/2004
  If CurrentQuery.FieldByName("GERARAUTOMATICO").AsString = "S" Then
    Dim sql As Object
    Set sql = NewQuery

    sql.Clear
    sql.Add("SELECT COUNT(1) AS QTDE FROM SFN_TIPODOCUMENTO WHERE GERARAUTOMATICO = :GERARAUTOMATICO AND HANDLE <> :HANDLE")
    sql.ParamByName("GERARAUTOMATICO").AsString = "S"
    sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.Active = True
    If sql.FieldByName("QTDE").AsInteger > 0 Then
      If VisibleMode Then
      	Beep
      End If
      bsShowMessage("Operação Ilegal! - Apenas um Tipo de Documento pode conter o Campo Documento Gerado pela Rotina Arquivo marcado.", "E")

      sql.Active = False
      Set sql = Nothing
      CanContinue = False

      Exit Sub
    End If
    sql.Active = False

    Set sql = Nothing
  End If

  'SMS 20961 - Kritian
  If CurrentQuery.FieldByName("ACEITAFATURACOMPLEMENTO").AsString = "N" Then
    If CurrentQuery.FieldByName("TABCONSIDERAGERACAOCOMPLEMENTO").AsInteger = 2 Then
      bsShowMessage("Para considerar a forma de geração do complemento da fatura é preciso marcar que o tipo de documento " + _
             "aceita faturas com complemento", "E")
      CanContinue = False
    End If
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOMODELODOCUMENTO" Then
		BOTAOMODELODOCUMENTO_OnClick
	End If
End Sub
