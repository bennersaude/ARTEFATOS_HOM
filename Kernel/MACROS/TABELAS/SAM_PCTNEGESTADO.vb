'HASH: 251C0205465C61B2249A7BC6DB86C9D8
'Macro: SAM_PCTNEGESTADO
'Mauricio Ibelli -04/05/2001 -sms 2226 -Selecionar grau validos
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub BOTAOGERARRELATORIO_OnClick()
If CurrentQuery.State <> 1 Then
	    bsShowMessage("O registro está em edição.","I")
    Else

    	Dim RelatorioHandle As Long
		Dim QueryBuscaHandleRelatorio As Object


		Set QueryBuscaHandleRelatorio=NewQuery

		QueryBuscaHandleRelatorio.Add("SELECT RELATORIOPACOTE FROM SAM_PARAMETROSPRESTADOR")
    	        QueryBuscaHandleRelatorio.Active=False
   		QueryBuscaHandleRelatorio.Active=True
   		RelatorioHandle=QueryBuscaHandleRelatorio.FieldByName("RELATORIOPACOTE").AsInteger

		If (RelatorioHandle = 0) Then
		 bsShowMessage("Relatório não está parametrizado","I")
		 CanContinue = False
		Else
		 ReportPreview(RelatorioHandle,"", False, False)
		End If

	    Set QueryBuscaHandleRelatorio=Nothing
	End If
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
    CurrentQuery.FieldByName("GRAUAGERAR").Clear
  End If
End Sub

Public Sub TABLE_AfterScroll()
 '-------------VALOR TOTAL DO PACOTE ------------------------------------------------
  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    Dim Interface As Object
    Dim valorpacte As Currency

    Set Interface = CreateBennerObject("BSPRE001.Rotinas")
    valorpacte = Interface.ValorTotalPacote(CurrentSystem, "SAM_PCTNEGESTADO", CurrentQuery.FieldByName("HANDLE").AsInteger)
    VALORPACOTE.Text = "Valor total do pacote: R$ " + Format(valorpacte, "#,##0.00")
    Set Interface = Nothing
  Else
    VALORPACOTE.Text = " "
  End If
  '-------------VALOR TOTAL DO PACOTE ------------------------------------------------
  
  If VisibleMode Then
  	If Not CurrentQuery.FieldByName("EVENTO").AsString = "" Then
    	GRAUAGERAR.LocalWhere = "ORIGEMVALOR = '7' AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @EVENTO)
  	End If
  Else
  	GRAUAGERAR.WebLocalWhere = "ORIGEMVALOR = '7' AND HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTO))
  	If WebMenuCode = "T5674" Then
  		EVENTO.ReadOnly = True
  	End If
  	If WebMenuCode = "T1303" Then
  		ESTADO.ReadOnly = True
  	End If
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  'Luciano T. Alberti - SMS 101345 - 21/08/2008 - Início
  Dim viRecord As Integer
  If RecordHandleOfTable("ESTADOS") > 0 Then
    viRecord = RecordHandleOfTable("ESTADOS")
  Else
    viRecord = CurrentQuery.FieldByName("ESTADO").AsInteger
  End If

  If viRecord > 0 Then
  'Luciano T. Alberti - SMS 101345 - 21/08/2008 - Fim
    If checkPermissao(CurrentSystem, CurrentUser, "E", viRecord, "E") = "N" Then
      bsShowMessage("Permissão negada. Usuário não pode excluir", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  'Luciano T. Alberti - SMS 101345 - 21/08/2008 - Início
  Dim viRecord As Integer
  If RecordHandleOfTable("ESTADOS") > 0 Then
    viRecord = RecordHandleOfTable("ESTADOS")
  Else
    viRecord = CurrentQuery.FieldByName("ESTADO").AsInteger
  End If

  If viRecord > 0 Then
  'Luciano T. Alberti - SMS 101345 - 21/08/2008 - Fim
    If checkPermissao(CurrentSystem, CurrentUser, "E", viRecord, "A") = "N" Then
      bsShowMessage("Permissão negada. Usuário não pode alterar", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If RecordHandleOfTable("ESTADOS") > 0 Then 'Luciano T. Alberti - SMS 101345 - 21/08/2008
    If checkPermissao(CurrentSystem, CurrentUser, "E", RecordHandleOfTable("ESTADOS"), "I") = "N" Then
      bsShowMessage("Permissão negada. Usuário não pode incluir", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Luciano T. Alberti - SMS 101345 - 21/08/2008 - Início
  Dim viRecord As Integer
  If RecordHandleOfTable("ESTADOS") > 0 Then
    viRecord = RecordHandleOfTable("ESTADOS")
  Else
    viRecord = CurrentQuery.FieldByName("ESTADO").AsInteger
  End If

  If viRecord > 0 Then
    If checkPermissao(CurrentSystem, CurrentUser, "E", viRecord, "I") = "N" Then
      bsShowMessage("Permissão negada. Usuário não pode incluir", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  'Luciano T. Alberti - SMS 101345 - 21/08/2008 - Fim

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
  Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
  Condicao = Condicao + " AND GRAUAGERAR = " + CurrentQuery.FieldByName("GRAUAGERAR").AsString

  If (CurrentQuery.FieldByName("ACOMODACAO").IsNull) Then
	Condicao = Condicao + " AND ACOMODACAO IS NULL "
  Else
    Condicao = Condicao + " AND ACOMODACAO = " + CurrentQuery.FieldByName("ACOMODACAO").AsString
  End If

  Linha = Interface.Vigencia(CurrentSystem, "SAM_PCTNEGESTADO", _
          "DATAINICIAL", _
          "DATAFINAL", _
          CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
          CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
          "ESTADO", _
          Condicao)

  If Linha <> "" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")
  SQL.Active = True

  If SQL.FieldByName("TOTAL").AsInteger = 1 Then
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")
    SQL.Active = True
    CurrentQuery.FieldByName("CONVENIO").Value = SQL.FieldByName("HANDLE").Value
  End If

  Set SQL = Nothing
End Sub

'MILANI -SMS -22609
Public Sub BOTAOINCLUIRITENS_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("BSPRE009.ROTINAS")
  Interface.ItensPacotes(CurrentSystem, "SAM_PCTNEGESTADO_GRAU", CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "BOTAOINCLUIRITENS" Then
    BOTAOINCLUIRITENS_OnClick
  End If
End Sub
