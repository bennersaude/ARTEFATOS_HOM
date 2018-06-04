'HASH: 61A8A3A2A7E3E41C1913DE6CC43CEB48
' macro: Z_GRUPOUSUARIOS_FILIAIS
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Q1 As Object
  Set Q1 = NewQuery

  Q1.Clear
  Q1.Add("SELECT FILIAL FROM Z_GRUPOUSUARIOS_FILIAIS WHERE USUARIO=" + CurrentQuery.FieldByName("USUARIO").AsString)
  Q1.Add("AND FILIAL=" + CurrentQuery.FieldByName("FILIAL").AsString)
  Q1.Add("AND HANDLE<>" + CurrentQuery.FieldByName("HANDLE").AsString)
  Q1.Active = True

  If Q1.FieldByName("FILIAL").AsInteger = CurrentQuery.FieldByName("FILIAL").AsInteger Then
    bsShowMessage("A Filial de Acesso já cadastrada.", "E")
    CanContinue = False
  End If

  Set Q1 = Nothing


End Sub

