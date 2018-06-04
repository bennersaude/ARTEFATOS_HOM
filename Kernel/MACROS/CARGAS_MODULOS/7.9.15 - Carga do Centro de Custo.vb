'HASH: 6E7E64D41327B68AE04271BD9A4B1B39
 

Public Sub BOTAOEXPORTARCC_OnClick()
Dim interface As Object
Set interface =CreateBennerObject("rotarq.rotinas")
interface.ExportaCorporativo(CurrentSystem,2)

End Sub
