'HASH: AA372D1E17072D256C55CA97474B1E8E

'#Uses "*ProcuraEvento"
'#Uses "*ProcuraTabelaGenerica"
'#Uses "*bsShowMessage"

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim Interface As Object
	Dim vColunas, vCriterio, vCampos, vTabela As String
	Dim ProcuraEvento As Long
	Set Interface = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
	End If

	vCriterio = " MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString
	vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos  ", True, "")

	If vHandle > 0 Then
	    CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
	End If

	Set Interface = Nothing
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim Interface As Object
	Dim vColunas, vCriterio, vCampos, vTabela As String
	Dim ProcuraEvento As Long
	Set Interface = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
	End If

	If CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Evento Inicial antes de selecionar Evento Final", "E")
	  Exit Sub
	End If

	vCriterio = " MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString
	vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos ", True, "")

	If vHandle > 0 Then
		CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
	End If

	Set Interface = Nothing

End Sub

Public Sub MASCARATGE_OnChange()
	If VisibleMode Then
		CurrentQuery.FieldByName("EVENTOINICIAL").Clear
		CurrentQuery.FieldByName("EVENTOFINAL").Clear
	End If
End Sub

Public Sub MASCARATGE_OnEnter()

End Sub

Public Sub TABELAPRECO_OnPopup(ShowPopup As Boolean)
'  If Len(EVENTO.Text)=0 Then
      Dim vHandle As Long
      ShowPopup =False
      vHandle =ProcuraTabelaGenerica(TABELAPRECO.Text)
      If vHandle <>0 Then
         CurrentQuery.Edit
         CurrentQuery.FieldByName("TABELAPRECO").Value =vHandle
      End If
'  End If
End Sub

Public Function CHECARPRECOREEMBOLSO()
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim SQL As Object

  CHECARPRECOREEMBOLSO =True

  Condicao =" AND DATAFINAL >=  DATAINICIAL "
  Condicao =Condicao +" AND HANDLE <> " +CurrentQuery.FieldByName("HANDLE").AsString

  Set Interface =CreateBennerObject("SAMGERAL.Vigencia")

  If CurrentQuery.FieldByName("ESTADO").IsNull Then
    Condicao =Condicao +" AND ESTADO IS NULL"
  Else
    Condicao =Condicao +" AND ESTADO = " +CurrentQuery.FieldByName("ESTADO").AsString
  End If

  If CurrentQuery.FieldByName("MUNICIPIO").IsNull Then
    Condicao =Condicao +" AND MUNICIPIO IS NULL"
  Else
    Condicao =Condicao +" AND MUNICIPIO = " +CurrentQuery.FieldByName("MUNICIPIO").AsString
  End If

  If VisibleMode Then
    Linha =Interface.Vigencia(CurrentSystem,"SAM_CONTRATO_PRECOREEMBOLSO","DATAINICIAL","DATAFINAL",CurrentQuery.FieldByName("DATAINICIAL").AsDateTime,CurrentQuery.FieldByName("DATAFINAL").AsDateTime,"CONTRATO",Condicao)
  End If

  If Linha ="" Then
    CHECARPRECOREEMBOLSO =False
  Else
    CHECARPRECOREEMBOLSO =True
    bsShowMessage(Linha, "E")
  End If

  Set Interface =Nothing

End Function

Public Sub TABLE_AfterScroll()
	If WebMode Then
		EVENTOINICIAL.WebLocalWhere = " A.MASCARATGE = @CAMPO(MASCARATGE)"
		EVENTOFINAL.WebLocalWhere = " A.MASCARATGE = @CAMPO(MASCARATGE)"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


	'Gabriel SMS 116237 - Inicio
    Dim qVerifica As Object
    Set qVerifica = NewQuery

    qVerifica.Add("SELECT MASCARATGE")
    qVerifica.Add("  FROM SAM_TGE")
	qVerifica.Add(" WHERE HANDLE = :HANDLE")

	qVerifica.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
    qVerifica.Active = True

	If qVerifica.FieldByName("MASCARATGE").AsInteger <> CurrentQuery.FieldByName("MASCARATGE").AsInteger Then
		CanContinue = False
		bsShowMessage("O Evento Inicial possui máscara diferente da informada. Altere o evento ou a máscara.", "E")
		Exit Sub
	End If

    qVerifica.Active = False
	qVerifica.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
	qVerifica.Active = True

	If qVerifica.FieldByName("MASCARATGE").AsInteger <> CurrentQuery.FieldByName("MASCARATGE").AsInteger Then
		CanContinue = False
		bsShowMessage("O Evento Final possui máscara diferente da informada. Altere o evento ou a máscara.", "E")
		Exit Sub
	End If
	'Gabriel SMS 116237 - Fim

  Dim SQLE As Object
  Set SQLE =NewQuery

 ' If CHECARPRECOREEMBOLSO Then //SMS 83756 - Willian - 28/6/2007 - Código comentado pois logo abaixo é realizado o mesmo procedimento
 '   CanContinue =False         // com a verificação da faixa de evento, o que não é realizado em CHECARPRECOREEMBOLSO.
 '   If VisibleMode Then
 '     RefreshNodesWithTable("SAM_CONTRATO_PRECOREEMBOLSO")
 '   End If
 '   Exit Sub
 ' End If

' Atribuir ESTRUTURAINICIAL E FINAL
  Dim SQLTGE,SQLMASC As Object
  Dim Estrutura As String

' Atribuir ESTRUTURAINICIAL
  Set SQLTGE =NewQuery
  SQLTGE.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")
  SQLTGE.ParamByName("HEVENTO").Value =CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
  SQLTGE.Active =True
  CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value =SQLTGE.FieldByName("ESTRUTURA").Value

' Atribuir ESTRUTURAFINAL
  SQLTGE.Active =False
  SQLTGE.ParamByName("HEVENTO").Value =CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
  SQLTGE.Active =True
  Estrutura =SQLTGE.FieldByName("ESTRUTURA").Value
  SQLTGE.Active =False
  Set SQLTGE =Nothing

  CurrentQuery.FieldByName("ESTRUTURAFINAL").Value =Estrutura

  If VisibleMode Then
    If (REGIMEATENDIMENTO.Visible And CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull) Then
      MsgBox("O campo 'Regime de Atendimento' deve ser preenchido.")
      CanContinue = False
    End If
  End If

  If CanContinue =True Then
  ' Checar Vigencia
    EstruturaI =CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString
    EstruturaF =CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString

    Set Interface =CreateBennerObject("SAMGERAL.Vigencia")

    Condicao =" CONTRATO = " +CurrentQuery.FieldByName("CONTRATO").AsString
    Condicao =Condicao +" AND HANDLE <> " +CurrentQuery.FieldByName("HANDLE").AsString
    condicao = condicao + " AND MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString

    If CurrentQuery.FieldByName("ESTADO").IsNull Then
      Condicao =Condicao +" AND ESTADO IS NULL"
    Else
      Condicao =Condicao +" AND ESTADO = " +CurrentQuery.FieldByName("ESTADO").AsString
    End If

    If CurrentQuery.FieldByName("MUNICIPIO").IsNull Then
      Condicao =Condicao +" AND MUNICIPIO IS NULL"
    Else
      Condicao =Condicao +" AND MUNICIPIO = " +CurrentQuery.FieldByName("MUNICIPIO").AsString
    End If

    If CurrentQuery.FieldByName("REGIMEATENDIMENTO").IsNull Then
      Condicao =Condicao +" AND REGIMEATENDIMENTO IS NULL"
    Else
      Condicao =Condicao +" AND REGIMEATENDIMENTO = " +CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsString
    End If

    Condicao = Condicao + " AND MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString

    Linha =Interface.EventoFx(CurrentSystem,"SAM_CONTRATO_PRECOREEMBOLSO",CurrentQuery.FieldByName("DATAINICIAL").AsDateTime,CurrentQuery.FieldByName("DATAFINAL").AsDateTime,EstruturaI,EstruturaF,Condicao)

    If Linha ="" Then
      CanContinue =True
    Else
      CanContinue =False
      bsShowMessage(Linha, "E")
      Exit Sub
    End If
    Set Interface =Nothing
  End If

  If CanContinue =True Then
    CanContinue =CheckEventosFx
  End If


End Sub


Public Function CheckEventosFx As Boolean
  CheckEventosFx =True
  If Not CurrentQuery.FieldByName("EVENTOINICIAL").IsNull Then
    If CurrentQuery.FieldByName("EVENTOFINAL").IsNull Then
      CurrentQuery.FieldByName("EVENTOFINAL").Value =CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
    Else
      If CurrentQuery.FieldByName("EVENTOINICIAL").Value <>CurrentQuery.FieldByName("EVENTOFINAL").Value Then
        Dim SQLI,SQLF As Object
        Set SQLI =NewQuery
        SQLI.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOI")
        SQLI.ParamByName("HEVENTOI").Value =CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
        SQLI.Active =True

        Set SQLF =NewQuery
        SQLF.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOF")
        SQLF.ParamByName("HEVENTOF").Value =CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
        SQLF.Active =True

        If SQLF.FieldByName("ESTRUTURA").Value <SQLI.FieldByName("ESTRUTURA").Value Then
          bsShowMessage("Evento final não pode ser menor que o evento inicial!", "E")
          If VisibleMode Then
            EVENTOFINAL.SetFocus
          End If
          CheckEventosFx =False
        End If
        Set SQLI =Nothing
        Set SQLF =Nothing
      End If
    End If
  End If
End Function
