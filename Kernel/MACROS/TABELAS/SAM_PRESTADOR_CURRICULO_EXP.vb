'HASH: A664662EB05E73C49A331ECBDE0F8853
'Macro: SAM_PRESTADOR_CURRICULO_EXP
'#Uses "*bsShowMessage"

Option Explicit

Dim NaoFaturarGuiasAnterior As String

Public Sub CEP_OnPopup(ShowPopup As Boolean)
	' Joldemar Moreira 12/06/2003
	' SMS 16059
	Dim vHandle As String
	Dim interface As Object

	ShowPopup = False

	Set interface = CreateBennerObject("ProcuraCEP.Rotinas")

	interface.Exec(CurrentSystem, vHandle)

	If vHandle <>"" Then
		Dim SQL As Object
		Set SQL = NewQuery

		SQL.Add("SELECT CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO,COMPLEMENTO   ")
		SQL.Add("  FROM LOGRADOUROS      ")
		SQL.Add(" WHERE CEP = :HANDLE ")

		SQL.ParamByName("HANDLE").Value = vHandle
		SQL.Active = True

		CurrentQuery.Edit
		CurrentQuery.FieldByName("CEP").Value = SQL.FieldByName("CEP").AsString
		CurrentQuery.FieldByName("ESTADO").Value = SQL.FieldByName("ESTADO").AsString
		CurrentQuery.FieldByName("MUNICIPIO").Value = SQL.FieldByName("MUNICIPIO").AsString
		CurrentQuery.FieldByName("BAIRRO").Value = SQL.FieldByName("BAIRRO").AsString
		CurrentQuery.FieldByName("LOGRADOURO").Value = SQL.FieldByName("LOGRADOURO").AsString
		CurrentQuery.FieldByName("LOGRADOUROCOMPLEMENTO").Value = SQL.FieldByName("COMPLEMENTO").AsString
	End If

	Set interface = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String
	Dim SQL As Object

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Set SQL = NewQuery

	SQL.Add("SELECT FISICAJURIDICA FROM SAM_PRESTADOR WHERE HANDLE = :P")

	SQL.ParamByName("P").Value = RecordHandleOfTable("SAM_PRESTADOR")
	SQL.Active = True

	If SQL.FieldByName("FISICAJURIDICA").AsInteger <>1 Then
		CanContinue = False
		bsShowMessage("O registro do currículo destina-se somenente a prestadores pessoa física", "E")
	End If

	Set SQL = Nothing
End Sub
