'HASH: 7DD136D7F72D4924C541093C7E6A80C1
'MACRO = TV_FORM_CA_PRESTADOR_ALTERACAO
Option Explicit

Public Sub TABLE_NewRecord()
	Dim vHandle   As Long
	Dim vResultado As String

	vHandle  = CLng(SessionVar("hAlteracoes"))

    Dim dll As Object

  	Set dll = CreateBennerObject("CA032.ALTERAPRESTADOR")
  	dll.ExibirAlteracoes(CurrentSystem, vHandle, vResultado)

	CurrentQuery.FieldByName("ALTERACOESPRESTADOR").AsString = vResultado

	Set dll = Nothing

End Sub
