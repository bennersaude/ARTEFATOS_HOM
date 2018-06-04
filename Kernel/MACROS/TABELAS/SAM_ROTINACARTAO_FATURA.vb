'HASH: 745C79DE07A5C1EEBE734AF43941CE52
'Macro: SAM_ROTINACARTAO_FATURA
'A funcao NodeInternalCode é utilizada para determinar se a carga correspondente é da Tarefas de Modelo,
'sendo, mostra o Tab - Modelo para agendamento, não sendo, mostra o Tab - Rotina
'Alteração: 26/12/2005
'      SMS: 52120 - Marcelo Barbosa
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
  If (VisibleMode) Then

    If (NodeInternalCode = 501) Then
      TABTIPOROTINA.Pages(0).Visible = False
      TABTIPOROTINA.Pages(1).Visible = True
    Else
      TABTIPOROTINA.Pages(0).Visible = True
      TABTIPOROTINA.Pages(1).Visible = False
    End If

  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  CanContinue = False

  Dim S As Object

  If ( VisibleMode And NodeInternalCode <> 501) Or (WebMode And WebMenuCode <> "T5140") Then

    Set S = NewQuery

    S.Add("SELECT COUNT(R.HANDLE) PARAMETROS")
    S.Add("  FROM SAM_ROTINACARTAO_CARTAO R, ")
    S.Add("       SAM_ROTINACARTAO_FATURA C")
    S.Add(" WHERE R.ROTINACARTAO = :HROTINACARTAO")
    S.Add("   AND C.ROTINACARTAO = R.HANDLE")
    S.Add("   AND C.SITUACAO = 'A'")

    S.ParamByName("HROTINACARTAO").Value = RecordHandleOfTable("SAM_ROTINACARTAO")
    S.Active = True


    If S.FieldByName("PARAMETROS").Value > 0 Then
      bsShowMessage("Há Parâmetros para Faturamento em Aberto.", "E")
      Exit Sub
    End If
  End If

  CanContinue = True

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  CanContinue = False

  'If CurrentQuery.State <> 1 Then
  '   MsgBox("Os parâmetros não podem estar em edição.")
  '   Exit Sub
  'End If

If ( VisibleMode And NodeInternalCode <> 501) Or (WebMode And WebMenuCode <> "T5140") Then

  If CurrentQuery.FieldByName("EMISSAOFATURA").AsDateTime < ServerDate Then
    bsShowMessage("Data emissão Inválida.", "E")
    EMISSAOFATURA.SetFocus
    Exit Sub
  End If

  If CurrentQuery.FieldByName("ORIGEMVENCIMENTO").AsInteger = 1 Then
    If CurrentQuery.FieldByName("DATAVENCIMENTO").IsNull Then
      bsShowMessage("Data de vencimento Inválida.", "E")
      DATAVENCIMENTO.SetFocus
      Exit Sub
    End If
  Else

    If CurrentQuery.FieldByName("ORIGEMVENCIMENTO").AsInteger = 2 Then
      If CurrentQuery.FieldByName("VENCIMENTOFATURA").IsNull Then
        bsShowMessage("Data emissão Inválida.", "E")
        VENCIMENTOFATURA.SetFocus
        Exit Sub
      End If
    Else
      bsShowMessage("Deve ser escolhida a origem do vencimento. ", "E")
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("TABDATACONTABIL").AsInteger = 1 Then
    If CurrentQuery.FieldByName("DATACONTABIL").IsNull Then
      bsShowMessage("Data Contábil não pode ser nula.", "E")
      DATACONTABIL.SetFocus
      Exit Sub
    End If
  Else
    If CurrentQuery.FieldByName("TABDATACONTABIL").AsInteger <> 2 Then
      bsShowMessage("Deve ser escolhida a origem do vencimento. ", "E")
      Exit Sub
    End If

  End If

Else
  'Macro Herdada da AGE_ROTINACARTAO_FATURA
  If CurrentQuery.FieldByName("DIASEMISSAOFATURA").AsInteger < 0 Then
    bsShowMessage("Data emissão Inválida.", "I")
    DIASEMISSAOFATURA.SetFocus
    Exit Sub
  End If

  If CurrentQuery.FieldByName("ORIGEMVENCIMENTOMODELO").AsInteger = 1 Then
    If CurrentQuery.FieldByName("DIASDATAVENCIMENTO").IsNull Then
      bsShowMessage("Data de vencimento Inválida.", "E")
      DIASDATAVENCIMENTO.SetFocus
      Exit Sub
    End If
  Else
    If CurrentQuery.FieldByName("ORIGEMVENCIMENTOMODELO").AsInteger = 2 Then
      If CurrentQuery.FieldByName("MESESVENCIMENTOFATURA").IsNull Then
      	bsShowMessage("Competência de vencimento Inválida.", "E")
        MESESVENCIMENTOFATURA.SetFocus
        Exit Sub
      End If
    Else
      bsShowMessage("Deve ser escolhida a origem do vencimento. ", "E")
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("TABDATACONTABILMODELO").AsInteger = 1 Then
    If CurrentQuery.FieldByName("DIASDATACONTABIL").IsNull Then
      bsShowMessage("Data Contábil não pode ser nula.", "E")
      DIASDATACONTABIL.SetFocus
      Exit Sub
    End If
  Else
    If CurrentQuery.FieldByName("TABDATACONTABILMODELO").AsInteger <> 2 Then
      bsShowMessage("Deve ser escolhida a origem do vencimento. ", "E")
      Exit Sub
    End If

  End If

End If

CanContinue = True

End Sub

Public Sub TABLE_NewRecord()
  If VisibleMode Then
    If (NodeInternalCode = 501)Then
      CurrentQuery.FieldByName("TABTIPOROTINA").Value = 2
    Else
      CurrentQuery.FieldByName("TABTIPOROTINA").Value = 1
	End If
  Else
      CurrentQuery.FieldByName("TABTIPOROTINA").Value = 1
  End If
End Sub
