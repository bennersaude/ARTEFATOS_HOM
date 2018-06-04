'HASH: 3ACA23B282E0282C946847639ECA28B5
'SAM_PRESTADOR_HABILITACAO
'02/01/2001 - Alterado por Paulo Garcia Junior - liberacao para edição do registro atraves dos parametros gerais de prestador
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"
Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub TABLE_AfterScroll()
	If LiberaHabilitacao <> "" Then
		DATA.ReadOnly = True
		EVENTO.ReadOnly = True
		HABILITACAO.ReadOnly = True
		PRESTADOR.ReadOnly = True
		TEMPORARIO.ReadOnly = True
	Else
		DATA.ReadOnly = False
		EVENTO.ReadOnly = False
		HABILITACAO.ReadOnly = False
		PRESTADOR.ReadOnly = False
		TEMPORARIO.ReadOnly = False
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaHabilitacao

	If Msg<>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaHabilitacao

	If Msg<>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaHabilitacao

	If Msg<>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Function LiberaHabilitacao As String
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT EDITAHABILITACAO FROM SAM_PARAMETROSPRESTADOR")

	SQL.Active = True

	If SQL.FieldByName("EDITAHABILITACAO").AsString <> "S" Then
		LiberaHabilitacao = "Carga somente para leitura!"
	End If
End Function
