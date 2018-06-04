'HASH: 4319AF52127FC6433712FE5839A9C917
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim vData As String

    vData = CurrentQuery.FieldByName("ANOCALENDARIO").AsString

    If Not IsDate(vData) Then
      bsShowMessage("Ano Calendário inválido !", "E")
      CanContinue = False
    End If


End Sub
