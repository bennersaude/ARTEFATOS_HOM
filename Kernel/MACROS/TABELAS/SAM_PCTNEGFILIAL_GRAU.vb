'HASH: B3490EB6C8F659371EA08B6B8604EDCF
'Macro: SAM_PCTNEGFILIAL_GRAU
'#Uses "*bsShowMessage"

Public Sub CODIGOTABELATISS_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
  	Dim vHandle As Long
  	Dim vHandleAntes As Long
  	Dim vCampos As String
  	Dim vColunas As String
  	Dim vCriterio As String

  	vHandleAntes=CurrentQuery.FieldByName("CODIGOTABELATISS").AsInteger


	ShowPopup = False
  	Set interface = CreateBennerObject("Procura.Procurar")

	vColunas = "TIS_TABELAPRECO.CODIGO|TIS_TABELAPRECO.DESCRICAO"

  	vCampos = "Código|Descrição"

  	vCriterio=" TIS_TABELAPRECO.VERSAOTISS = (SELECT MAX(B.HANDLE) FROM TIS_VERSAO B WHERE B.ATIVODESKTOP = 'S') "

	vHandle = interface.Exec(CurrentSystem, "TIS_TABELAPRECO", vColunas, 1, vCampos, vCriterio, "Tabela Tiss", True, "")

  	If vHandle <>vHandleAntes Then
    	CurrentQuery.Edit
    	CurrentQuery.FieldByName("CODIGOTABELATISS").Value = vHandle
    	CurrentQuery.FieldByName("EVENTOAGERAR").Value = Null
    	CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
  	End If

  	Set interface = Nothing

  	If vHandle = 0 Then
    	CurrentQuery.FieldByName("CODIGOTABELATISS").Value = Null
    	CurrentQuery.FieldByName("EVENTOAGERAR").Value = Null
    	CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
  	End If

End Sub

Public Sub GRAUAGERAR_OnExit()
	Dim vHandle As Long
	Dim SQL As Object
	Set SQL = NewQuery
	SQL.Add("SELECT UTILIZAFILTROGRAUPACOTE FROM SAM_PARAMETROSPRESTADOR")
	SQL.Active = True

  	If SQL.FieldByName("UTILIZAFILTROGRAUPACOTE").AsString = "S" Then
	    SQL.Active = False
    	If Not CurrentQuery.FieldByName("GRAUAGERAR").IsNull Then
      		SQL.Clear
      		SQL.Add("SELECT HANDLE FROM SAM_GRAU WHERE ORIGEMVALOR <> '7' AND HANDLE = " + CurrentQuery.FieldByName("GRAUAGERAR").AsString)
      		SQL.Active = True
      		If SQL.EOF Then
        		bsShowMessage("Grau pacote não é válido!", "E")
        		CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
        		Set SQL = Nothing
        		Exit Sub
      		End If
    	End If
  	End If
  	Set SQL = Nothing

End Sub

Public Sub TABLE_AfterInsert()

	VerificaPacoteFinalizado(CurrentQuery.FieldByName("PACOTE").AsInteger)

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String
    If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	If (VerificaPacoteFinalizado(CurrentQuery.FieldByName("PACOTE").AsInteger))Then
	  CanContinue = False
	  Exit Sub
	End If

End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
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

'#Uses "*ProcuraEvento"

Public Sub EVENTOAGERAR_OnPopup(ShowPopup As Boolean)
  ShowPopup = False
  If CurrentQuery.FieldByName("CODIGOTABELATISS").IsNull Then
    bsShowMessage("Informar uma tabela Tiss!", "I")
    Exit Sub
  End If

  Dim interface As Object
  Dim vHandle As Long
  Dim vHandleAntes As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCodigoTabelaTiss As String

  vCodigoTabelaTiss= CurrentQuery.FieldByName("CODIGOTABELATISS").AsString
  vHandleAntes=CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TGE.ESTRUTURA|SAM_TGE.DESCRICAO"

  vCampos = "Evento|Descrição"

  vCriterio=" SAM_TGE.HANDLE IN ( SELECT EVENTO FROM SAM_TGE_TABELATISS WHERE SAM_TGE_TABELATISS.TABELATISS="+vCodigoTabelaTiss+" )"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela de Eventos", True, "")

  If vHandle <>vHandleAntes Then
    CurrentQuery.Edit
    If vHandle = 0 Then vHandle=Null

    CurrentQuery.FieldByName("EVENTOAGERAR").Value = vHandle
    CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
  End If

  Set interface = Nothing

  If vHandle = 0 Then
    CurrentQuery.FieldByName("EVENTOAGERAR").Value = Null
    CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
  End If

End Sub


Public Sub CODIGOPAGTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PERCENTUALPGTO.CODIGOPAGTO|DESCRICAO|INCIDENCIAMINIMA|PERCENTUALPGTOINCIDENCIA1|PERCENTUALPGTODEMAIS|USADOAUTORIZACAO|USADOPAGAMENTO"

  vCampos = "Código|Descrição|Incidência Mínima|% Pagto Inc 1|% Pagto Demais|Usado Autorização|Usado Pagto"

  vHandle = interface.Exec(CurrentSystem, "SAM_PERCENTUALPGTO", vColunas, 1, vCampos, vCriterio, "Tabela de Códigos de Pagamentos", True, "")


  If vHandle <>0 Then
    CurrentQuery.Edit
    If vHandle = 0 Then vHandle=Null

    CurrentQuery.FieldByName("CODIGOPAGTO").Value = vHandle
  End If

  Set interface = Nothing

  If vHandle = 0 Then
    CurrentQuery.FieldByName("CODIGOPAGTO").Value = Null
  End If

End Sub


Public Sub GRAUAGERAR_OnPopup(ShowPopup As Boolean)
  Dim vHandleGrau As Long
  Dim interface As Object
  Dim SQL As Object
  Dim vColunas, vCriterio, vCampos, vTabela As String

  ShowPopup = False

  If CurrentQuery.FieldByName("EVENTOAGERAR").IsNull Then
    bsShowMessage("Informar evento !","I")
    Exit Sub
  End If

  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_GRAU.GRAU|SAM_GRAU.Z_DESCRICAO|SAM_TIPOGRAU.DESCRICAO"
  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")
  SQL.Active = True
  If SQL.FieldByName("UTILIZAFILTROGRAUPACOTE").AsString = "S" Then
    vCriterio = " ORIGEMVALOR <> '7' "
  Else
    vCriterio = ""
  End If

  If vCriterio <>"" Then
    vCriterio = vCriterio + " AND "
  End If
  vCriterio = vCriterio + "(SAM_GRAU.HANDLE IN (SELECT S.HANDLE                          "
  vCriterio = vCriterio + "                      FROM SAM_GRAU S                         "
  vCriterio = vCriterio + "                     WHERE (S.VERIFICAGRAUSVALIDOS IS NULL OR S.VERIFICAGRAUSVALIDOS = 'N') "
  vCriterio = vCriterio + "                   )                                          "
  vCriterio = vCriterio + "OR                                                            "
  vCriterio = vCriterio + "SAM_GRAU.HANDLE IN (SELECT S.HANDLE                           "
  vCriterio = vCriterio + "                      FROM SAM_GRAU S                         "
  vCriterio = vCriterio + "                     WHERE S.HANDLE IN (SELECT T.GRAU         "
  vCriterio = vCriterio + "                                          FROM SAM_TGE_GRAU T "
  vCriterio = vCriterio + "                                         WHERE T.EVENTO = " + CurrentQuery.FieldByName("EVENTOAGERAR").AsString
  vCriterio = vCriterio + "                                       )                      "
  vCriterio = vCriterio + "                   ))                                         "

  vCampos = "Código do Grau|Descrição|Tipo do Grau"
  vTabela = "SAM_GRAU|SAM_TIPOGRAU[SAM_TIPOGRAU.HANDLE = SAM_GRAU.TIPOGRAU]"

  vHandleGrau = interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Tabela de graus ", True, "")

  If vHandleGrau >0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAUAGERAR").Value = vHandleGrau
  End If

  Set interface = Nothing

  If vHandleGrau = 0 Then
    CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
  End If

  Set SQL = Nothing

End Sub

Public Sub TABELAUSVALOR_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "DESCRICAO|DATAINICIAL|DATAFINAL|VALORUSHONORARIO|VALORUSCUSTOOPERACIONAL"

  vCriterio = ""
  vCampos = "Descrição da Tabela|Data Inicial|Data Final|Vr. US Honorário|Vr US Custo Operac"

  vHandle = interface.Exec(CurrentSystem, "SAM_TABUS|SAM_TABUS_VLR[SAM_TABUS.HANDLE = SAM_TABUS_VLR.TABELAUS]", vColunas, 1, vCampos, vCriterio, "Tabela de US", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABELAUSVALOR").Value = vHandle
  End If

  Set interface = Nothing

  If vHandle = 0 Then
    CurrentQuery.FieldByName("TABELAUSVALOR").Value = Null
  End If


End Sub



Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim Msg As String
    If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	If (VerificaPacoteFinalizado(CurrentQuery.FieldByName("PACOTE").AsInteger))Then
	  CanContinue = False
	  Exit Sub
	End If

	Dim vHandle As Long
	Dim SQL As Object
	Set SQL = NewQuery
	SQL.Add("SELECT UTILIZAFILTROGRAUPACOTE FROM SAM_PARAMETROSPRESTADOR")
	SQL.Active = True

	If SQL.FieldByName("UTILIZAFILTROGRAUPACOTE").AsString = "S" Then
  		SQL.Active = False
  		If Not CurrentQuery.FieldByName("GRAUAGERAR").IsNull Then
    		SQL.Clear
    		SQL.Add("SELECT HANDLE FROM SAM_GRAU WHERE ORIGEMVALOR <> '7' AND HANDLE = " + CurrentQuery.FieldByName("GRAUAGERAR").AsString)
    		SQL.Active = True
    		If SQL.EOF Then
      			bsShowMessage("Grau pacote não é válido!", "E")
      			CanContinue = False
      			Set SQL = Nothing
      			Exit Sub
    		End If
		End If
	End If

	Set SQL = Nothing

End Sub

Public Function VerificaPacoteFinalizado(HandlePacote As Long) As Boolean
	Dim SQL
	Set SQL = NewQuery
	SQL.Add("SELECT DATAFINAL FROM SAM_PCTNEGFILIAL  WHERE HANDLE = :HPCTNEGFILIAL")
	SQL.ParamByName("HPCTNEGFILIAL").AsInteger = HandlePacote
	SQL.Active = True

	If((Not SQL.FieldByName("DATAFINAL").IsNull)And(SQL.FieldByName("DATAFINAL").AsDateTime <= ServerDate))Then
		bsShowMessage("Pacote finalizado não permite manutenções", "E")
		VerificaPacoteFinalizado = True
	Else
		VerificaPacoteFinalizado = False
	End If

	Set SQL = Nothing
End Function
