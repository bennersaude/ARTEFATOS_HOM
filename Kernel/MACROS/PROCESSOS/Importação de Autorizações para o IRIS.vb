'HASH: F8A355AB4AB0BD11543A30ED6CE3D8E6
Public Sub Main
  Dim Interface As Object
  Set Interface = CreateBennerObject("Benner.Saude.Implementation.IntegracaoIris.CImporta")
  Interface.Importa(CurrentSystem, FormatDateTime2("dd-mm-yyyy",ServerDate))
  Set Interface = Nothing
End Sub
