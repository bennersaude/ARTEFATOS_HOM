'HASH: BBD057561B6CADD29DCC274C26FDD221
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"

Public Function ProcuraGrau ( ) As Long
	Dim Interface As Object
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim qParamAtend As Object
	Set qParamAtend = NewQuery

	qParamAtend.Add("SELECT FILTRARGRAUSVALIDOS FROM PARAMETROSATENDIMENTO")

	Set qParamAtend.Active = True
	Set Interface = CreateBennerObject("Procura.Procurar")

	vColunas = "SAM_GRAU.GRAU|SAM_GRAU.Z_DESCRICAO|SAM_TIPOGRAU.DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

	'-----------------------------------------------------------------------------------
	If qParamAtend.FieldByName("VERIFICAGRAUSVALIDOS").AsString = "S" Then
		vCriterio = "HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE HANDLE = " + CurrentQuery.FieldByName("EVENTO").AsString + ")"
	Else
		vCriterio = ""
	End If
	'------------------------------------------------------------------------------------

	vCampos = "Código do Grau|Descrição|Tipo do Grau|Graus Válidos"
	vTabela = "SAM_GRAU|SAM_TIPOGRAU[SAM_TIPOGRAU.HANDLE = SAM_GRAU.TIPOGRAU]"
	ProcuraGrau = Interface.Exec(CurrentUser, vTabela, vColunas, 2, vCampos, vCriterio, "Graus de Atuação", True, GRAU.Text)

	Set Interface = Nothing
End Function

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)

  If vHandle <> 0 Then
    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If

    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	'Inicio SMS 37619 Wagner Santos 20/07/2005
	ShowPopup = False

	Dim vHandle As Long
	Dim vColunas As String
	Dim vCriterio As String
	Dim vCampos As String
	Dim ProcuraDll As Object
	Dim SQL As Object
	Dim vGrau As String
	Dim vContador As Integer
	Dim vColunaIndice As Integer

	vCriterio = ""
	vColunaIndice = 1
	vGrau = TiraAcento(GRAU.Text, True)

	For vContador = 1 To Len(GRAU.Text)
		If InStr(GRAU.Text, "0") Or _
			 InStr(GRAU.Text, "1") Or _
			 InStr(GRAU.Text, "2") Or _
			 InStr(GRAU.Text, "3") Or _
			 InStr(GRAU.Text, "4") Or _
			 InStr(GRAU.Text, "5") Or _
			 InStr(GRAU.Text, "6") Or _
			 InStr(GRAU.Text, "7") Or _
			 InStr(GRAU.Text, "8") Or _
			 InStr(GRAU.Text, "9") Then
			vColunaIndice = 1
		Else
			vColunaIndice = 2
			Exit For
		End If
	Next

	Set SQL = NewQuery

	SQL.Add("SELECT GRAUSVALIDOSNOALERTA FROM SAM_PARAMETROSPRESTADOR")

	SQL.Active = True

	Set ProcuraDll = CreateBennerObject("PROCURA.PROCURAR")

	vColunas = "SAM_GRAU.GRAU|SAM_GRAU.DESCRICAO|SAM_TIPOGRAU.DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

	If GRAU.Text = "" Then
		If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
			If SQL.FieldByName("GRAUSVALIDOSNOALERTA").AsString = "S" Then
				vCriterio = "SAM_GRAU.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + ")"
			Else
				vCriterio = ""
			End If
		Else
			bsShowMessage("Informar evento!", "I")
			EVENTO.SetFocus
			Exit Sub
		End If
	Else
		If vColunaIndice = 1 Then
			vCriterio = "GRAU = " + vGrau
		Else
			vCriterio = "Z_DESCRICAO LIKE '" + vGrau + "%'"
			vGrau = ""
		End If
	End If

  vCampos = "Código do Grau|Descrição|Tipo do Grau|Graus Válidos"
  vHandle = ProcuraDll.Exec(CurrentSystem, "SAM_GRAU|SAM_TIPOGRAU[SAM_TIPOGRAU.HANDLE = SAM_GRAU.TIPOGRAU]", vColunas, vColunaIndice, vCampos, vCriterio, "Graus de atuação", True, vGrau)

  If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If

  Set ProcuraDll = Nothing
  Set SQL = Nothing
  'Fim SMS 37619
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
	  	Dim SQL As Object

		Set SQL = NewQuery
  		SQL.Add("SELECT GRAUSVALIDOSNOALERTA FROM SAM_PARAMETROSPRESTADOR")
	  	SQL.Active = True

		If SQL.FieldByName("GRAUSVALIDOSNOALERTA").AsString = "S" Then
  			GRAU.WebLocalWhere = "A.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTO))"
		Else
  			GRAU.WebLocalWhere = ""
		End If
	End If
End Sub
