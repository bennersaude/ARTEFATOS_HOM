'HASH: F91587FE91B5AB3AE55EECA5CD95DD9C
'MACRO: TV_FORM_COMPETFIN
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim  obj As Object
	Dim retorno As String

    Set obj = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.SfnCompetfinBLL, Benner.Saude.Financeiro.Business")

    obj.ClearParameters
    obj.AddParameter(pdtDateTime, CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
    obj.AddParameter(pdtInteger, CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)

	retorno = obj.Execute("InserirCompetencia")

	bsShowMessage(retorno, "I")

End Sub
