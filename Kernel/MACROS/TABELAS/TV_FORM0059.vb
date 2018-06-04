'HASH: 75EFF46CFD0DC7A0563BD1B3CA1F493F
Dim Eventos As String
Dim Graus As String
Dim hIncomp As Double
Dim Interface As Object


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Eventos = CurrentQuery.FieldByName("EVENTO").AsString
  Graus = CurrentQuery.FieldByName("GRAU").AsString
  hIncomp = RecordHandleOfTable("SAM_INCOMP_EVENTOS_GERAL")

Set Interface = CreateBennerObject("SAMDUPEVENTOS.DuplicarIncompatibilidade")
Set Interface.Exec(CurrentSystem, Eventos, Graus, hIncomp, "A")
End Sub

