'HASH: 66E9C544117871D03248BAA0CF9E6AD8
'Macro: SAM_BENEF_PRECOREEMBOLSO
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraTabelaGenerica"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
   ShowPopup = False
   PopUpEvento "EVENTOFINAL", EVENTOFINAL.Text
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
   ShowPopup = False
   PopUpEvento "EVENTOINICIAL", EVENTOINICIAL.Text
End Sub

Public Sub PopUpEvento(NomeCampoEvento As String, TextCampoEvento As String)
    Dim vHandle As Long
	Dim Interface As Object
	Dim vColunas, vCriterio, vCampos, vTabela As String
	Dim ProcuraEvento As Long
	Dim vNumeroColuna As Integer

	If CurrentQuery.FieldByName("MASCARATGE").AsInteger = 0 Then
	  bsShowMessage("Necessário escolher Máscara da TGE antes de selecionar eventos", "E")
	  Exit Sub
	End If

    Set Interface = CreateBennerObject("Procura.Procurar")
	vColunas = " SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"
	vCriterio = " SAM_TGE.ULTIMONIVEL = 'S' AND MASCARATGE = " + CurrentQuery.FieldByName("MASCARATGE").AsString
	vCampos = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
	vTabela = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"

	vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Eventos  ", False, TextCampoEvento)

	If vHandle > 0 Then
	    CurrentQuery.Edit
		CurrentQuery.FieldByName(NomeCampoEvento).Value = vHandle
	End If

	Set Interface = Nothing
End Sub
Public Sub MASCARATGE_OnChange()
	If VisibleMode Then
		CurrentQuery.FieldByName("EVENTOINICIAL").Clear
		CurrentQuery.FieldByName("EVENTOFINAL").Clear
	End If
End Sub
Public Sub TABELAPRECO_OnPopup(ShowPopup As Boolean)
    Dim vHandle As Long
    ShowPopup =False
    vHandle =ProcuraTabelaGenerica(TABELAPRECO.Text)
    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("TABELAPRECO").Value =vHandle
    End If
End Sub
Public Sub EVENTOINICIAL_OnChange()
	If VisibleMode Then
		CurrentQuery.FieldByName("ESTRUTURAINICIAL").Clear
	End If
End Sub
Public Sub EVENTOFINAL_OnChange()
	If VisibleMode Then
		CurrentQuery.FieldByName("ESTRUTURAFINAL").Clear
	End If
End Sub
