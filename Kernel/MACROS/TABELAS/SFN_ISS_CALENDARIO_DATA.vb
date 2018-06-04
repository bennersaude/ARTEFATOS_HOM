'HASH: 872C1453E7D10566FA8EB26C48772C8E
'#Uses "*bsShowMessage"


Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    CanContinue = False
    bsShowMessage("Data Inicial não pode ser maior que a Data Final", "E")
  End If
End Sub

