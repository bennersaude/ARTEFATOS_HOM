'HASH: 5CB0A9C70DE62843BBE7108031BC756D
'Macro: SAM_BENEF_PRECOREEMBDOT_GER
'#Uses "*bsShowMessage"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"

Option Explicit
Public Sub CBOSPESQUISA_OnChange()
	CurrentQuery.FieldByName("CBOS").Clear
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long
	Dim Interface As Object
	Dim vColunas, vCampos, vTabela, vCriterio As String

	Set Interface =CreateBennerObject("Procura.Procurar")

	ShowPopup = False

	vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"
	vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
	vCriterio = " SAM_TGE.ULTIMONIVEL = 'S' "

	vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos", False, EVENTO.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value =vHandle
	End If

	Set Interface = Nothing
End Sub

Public Sub CBOSPESQUISA_OnPopup(ShowPopup As Boolean)
    Dim Interface As Object
    Dim vHandle, vNumeroColuna As Long
    Dim vCampos As String
    Dim vColunas As String

    ShowPopup = False

    Set Interface = CreateBennerObject("Procura.Procurar")

    vColunas = "TIS_VERSAO.VERSAO|TIS_CBOS.CODIGO|TIS_CBOS.DESCRICAO"
    vCampos = "Versão TISS|Código do CBOS|Descrição do CBOS"

	If IsNumeric(CBOSPESQUISA.Text) Then
		vNumeroColuna = 2
	Else
		vNumeroColuna = 3
	End If

    vHandle = Interface.Exec(CurrentSystem,"TIS_CBOS|TIS_VERSAO[TIS_CBOS.VERSAOTISS = TIS_VERSAO.HANDLE]", vColunas, vNumeroColuna, vCampos, "", "CBOS", False, CBOSPESQUISA.Text)

    If (vHandle <> 0) Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("CBOSPESQUISA").Value = vHandle
	End If
	Set Interface = Nothing
End Sub

Public Sub TABELAFILME_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaFilme(TABELAFILME.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAFILME").Value = vHandle
	End If
End Sub

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaUS(TABELAUS.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAUS").Value = vHandle
	End If
End Sub
