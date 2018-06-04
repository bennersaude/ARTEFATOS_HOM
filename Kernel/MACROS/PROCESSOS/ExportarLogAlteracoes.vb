'HASH: 78986657178EE07A03D61F1AEEC2E08B
Option Explicit

Public Sub Main
	Dim fullExport As String

    fullExport = SessionVar("fullExport")'""

    If fullExport <> "S" Then
      fullExport = "N"
    End If

    Dim dll As CSBusinessComponent
	Set dll = BusinessComponent.CreateInstance("Benner.Saude.Conecta.Business.ExporterHelper, Benner.Saude.Conecta.Business")
	dll.AddParameter(pdtAutomatic, SessionVar("TYPE"))
	dll.AddParameter(pdtAutomatic, fullExport)
	dll.Execute("RunExporter")
	Set dll = Nothing
End Sub
