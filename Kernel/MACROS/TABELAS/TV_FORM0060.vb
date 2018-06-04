'HASH: 5BF4DD008E185390CBA38433AFBF9A29
Dim Eventos As String
Dim Graus As String
Dim hIncomp As Double
Dim Interface As Object


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Eventos = CurrentQuery.FieldByName("EVENTO").AsString
  Graus = CurrentQuery.FieldByName("GRAU").AsString
  hIncomp = RecordHandleOfTable("SAM_INCOMP_EVENTOS_GERAL")

Set Interface = CreateBennerObject("SAMDUPEVENTOS.DuplicarIncompatibilidade")
Set Interface.Exec(CurrentSystem, Eventos, Graus, hIncomp, "P")
End Sub
