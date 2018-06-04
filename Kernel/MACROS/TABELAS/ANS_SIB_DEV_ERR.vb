'HASH: B9538CC0BD4974E93EA520013F73DE25
'Macro: ANS_SIB_DEV_ERR

Public Function ConsultarDadosBeneficiario

	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Active = False
	SQL.Clear
	SQL.Add("SELECT COUNT(SC.HANDLE) QTDE                                       ")
	SQL.Add("  FROM ANS_SIB_ROTINA_ENVIO SRE                                    ")
	SQL.Add("  JOIN ANS_SIB_CADASTRO     SC  ON SC.HANDLE      = SRE.NOME       ")
	SQL.Add("  JOIN ANS_SIB_ROTINA       SR1 ON SR1.HANDLE     = SRE.ROTINASIB  ")
	SQL.Add("  JOIN ANS_SIB_DEV_ERR      SDE ON SDE.SEQUENCIAL = SRE.SEQUENCIAL ")
	SQL.Add("  JOIN ANS_SIB_ROTINA       SR2 ON SR2.HANDLE     = SDE.ROTINASIB  ")
	SQL.Add(" WHERE SDE.HANDLE      = :HANDLE                                   ")
	SQL.Add("   AND SR1.COMPETENCIA = SR2.COMPETENCIA                           ")
	SQL.Add("   AND SR1.OPERADORA   = SR2.OPERADORA                             ")
	SQL.Add("   AND SR1.TIPO        = 1                                         ")

    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL.Active = True

	If SQL.FieldByName("QTDE").AsInteger = 1 Then

		SQL.Active = False
		SQL.Clear
		SQL.Add("SELECT SC.NOME                                                     ")
		SQL.Add("  FROM ANS_SIB_ROTINA_ENVIO SRE                                    ")
		SQL.Add("  JOIN ANS_SIB_CADASTRO     SC  ON SC.HANDLE      = SRE.NOME       ")
		SQL.Add("  JOIN ANS_SIB_ROTINA       SR1 ON SR1.HANDLE     = SRE.ROTINASIB  ")
		SQL.Add("  JOIN ANS_SIB_DEV_ERR      SDE ON SDE.SEQUENCIAL = SRE.SEQUENCIAL ")
		SQL.Add("  JOIN ANS_SIB_ROTINA       SR2 ON SR2.HANDLE     = SDE.ROTINASIB  ")
		SQL.Add(" WHERE SDE.HANDLE      = :HANDLE                                   ")
		SQL.Add("   AND SR1.COMPETENCIA = SR2.COMPETENCIA                           ")
		SQL.Add("   AND SR1.OPERADORA   = SR2.OPERADORA                             ")
  		SQL.Add("   AND SR1.TIPO        = 1                                         ")

	    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		SQL.Active = True

		ConsultarDadosBeneficiario = "Beneficiário sugerido: " + SQL.FieldByName("NOME").AsString

	ElseIf SQL.FieldByName("QTDE").AsInteger > 1 Then
		ConsultarDadosBeneficiario = "Foi encontrado mais de um sequencial linha para a rotina remessa desta competência."
	Else
		ConsultarDadosBeneficiario = "Sequencial linha não encontrado na rotina remessa desta competência."
	End If

	SQL.Active = False
	Set SQL = Nothing

End Function

Public Sub TABLE_AfterScroll()
	ROTULODADOSBENEFICIARIO.Text = ConsultarDadosBeneficiario
End Sub
