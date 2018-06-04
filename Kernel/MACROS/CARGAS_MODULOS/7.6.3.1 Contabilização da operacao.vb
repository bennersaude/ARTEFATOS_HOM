'HASH: 99C1FC6758843733126E775EE8EB9FED
 

Public Sub EXPORTARHISTORICOS_OnClick()

'Dim var1 As Integer
'var1=1
Dim interface As Object
Set interface =CreateBennerObject("rotarq.rotinas")
interface.ExportaCorporativo(CurrentSystem,1)
Set interface =Nothing

'Public Sub BOTAOEXPORTAR_OnClick()

'Dim interface As Object
'Set interface=CreateBennerObject("rotarq.rotinas")
'interface.ExportarModelo(CurrentQuery.FieldByName("HANDLE").AsInteger)
'Set interface=Nothing


'End Sub

End Sub
