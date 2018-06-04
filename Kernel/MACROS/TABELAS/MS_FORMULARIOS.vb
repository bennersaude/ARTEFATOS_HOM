'HASH: C113552F5E6D1ABE3218979E6982826E


Dim Obj As Object

Public Sub DUPLICAR_OnClick()
  Set Obj = CreateBennerObject("BSMed001.GlobalMS")
  Obj.Exec(CurrentSystem, 1)

  Set Obj = Nothing
End Sub

