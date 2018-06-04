'HASH: 658328B5107F36E835536CFC2C92ECB2
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()

  'Luciano T. Alberti - SMS 95250 - 31/03/2008 - Início
  If WebMode Then

    If WebMenuCode = "T1505" Then 'Menu de especialidades
      ESPECIALIDADE.ReadOnly = True
    Else
      ESPECIALIDADE.ReadOnly = False
    End If

    If WebMenuCode = "T1515" Then 'Menu da Área do Livro
      AREALIVRO.ReadOnly = True
    Else
      AREALIVRO.ReadOnly = False
    End If

  End If
  'Luciano T. Alberti - SMS 95250 - 31/03/2008 - Fim

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT HANDLE")
	SQL.Add("FROM SAM_PRESTADOR_LIVRO")
	SQL.Add("WHERE AREA = :HAREA")
	SQL.Add("  AND ESPECIALIDADE = :HESPECIALIDADE")

	SQL.ParamByName("HAREA").Value = CurrentQuery.FieldByName("AREALIVRO").AsInteger
	SQL.ParamByName("HESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
	SQL.Active = True

	If Not SQL.EOF Then
		CanContinue = False
		bsShowMessage("Existe algum prestador com a especialidade cadastrada em seu livro", "I")
	End If

	Set SQL = Nothing
End Sub


