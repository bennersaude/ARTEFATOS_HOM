'HASH: 926E21C4D8FBD4F97D2591793B282C3C
 

Public Sub BOTAOEXPORTARCLASSES_OnClick()
Dim interface As Object
Set interface =CreateBennerObject("rotarq.rotinas")
interface.ExportaCorporativo(CurrentSystem,3)

End Sub
