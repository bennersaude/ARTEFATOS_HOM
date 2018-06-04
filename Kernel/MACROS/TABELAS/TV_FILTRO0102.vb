'HASH: 0D2DFF41EA28ABB60F0F043B07BDA778
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim interface As Object

Dim pdCompetencia As Date
Dim psContrato As String
Dim psFilial As String

pdCompetencia = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
psContrato = CurrentQuery.FieldByName("CONTRATO").AsString
psFilial = CurrentQuery.FieldByName("FILIAL").AsString

    Set interface =CreateBennerObject("SamPlanilha.UIRotinas")
    interface.GuiasNaoProcessadas(CurrentSystem, pdCompetencia, psContrato, psFilial)
    Set interface =Nothing

End Sub
