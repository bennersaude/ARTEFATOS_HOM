'HASH: E22F4A89BE902B035EFCD0A99C9C213A
Sub main()

Dim Obj As Object
Set Obj = CreateBennerObject("IRIS.Rotinas")
Dim resultado As String
resultado = Obj.ExportaXML_IRIS(CurrentSystem, CStr(SessionVar(pCaminho)),ServerDate-1, ServerDate)
Set Obj = Nothing

End Sub
