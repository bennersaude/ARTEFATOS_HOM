'HASH: 3F0CC8B41533B020C13145AED32871DA
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
  Dim Qtmp As Object

  Set Qtmp = NewQuery

  Qtmp.Active = False
  Qtmp.Clear
  Qtmp.Add("SELECT USARCALENDARIODN FROM SAM_PARAMETROSPROCCONTAS ")
  Qtmp.Active = True

  If (WebMode) Then
    USARCALENDARIODN.ReadOnly = (Qtmp.FieldByName("USARCALENDARIODN").AsInteger = 2)
  Else
    USARCALENDARIODN.Visible = (Qtmp.FieldByName("USARCALENDARIODN").AsInteger = 1)
  End If

  If (WebMode) Then
      MODELOGUIA.WebLocalWhere    = " (HANDLE In (Select HANDLE " + _
                               "               FROM SAM_TIPOGUIA_MDGUIA " + _
                               "              WHERE REGIMEPAGTO = 'R') " + _
                               "  And @CAMPO(TABREGIMEPGTO) = 2) " + _
                               "  Or " + _
                               " ((HANDLE In (Select HANDLE " + _
                               "               FROM SAM_TIPOGUIA_MDGUIA) " + _
                               " And @CAMPO(TABREGIMEPGTO) = 3)) " + _
                               " Or ((@CAMPO(TABREGIMEPGTO) = 1) AND HANDLE IN (SELECT MD.HANDLE " + _
                               "               FROM SAM_TIPOGUIA_MDGUIA MD " + _
						       "                    JOIN SAM_TIPOGUIA TG ON (TG.HANDLE = MD.TIPOGUIA) " + _
                               "              WHERE MD.REGIMEPAGTO = 'C' AND TG.TIPOGUIATISS = @CAMPO(TIPOPEGTISS)) " + _
                               "  )"


  Else
    MODELOGUIA.LocalWhere    = " (HANDLE In (Select HANDLE " + _
                               "               FROM SAM_TIPOGUIA_MDGUIA " + _
                               "              WHERE REGIMEPAGTO = 'R') " + _
                               "  And @TABREGIMEPGTO = 2) " + _
                               "  Or " + _
                               " ((HANDLE In (Select HANDLE " + _
                               "               FROM SAM_TIPOGUIA_MDGUIA) " + _
                               " And @TABREGIMEPGTO = 3)) " + _
                               " Or ((@TABREGIMEPGTO = 1) AND HANDLE IN (SELECT MD.HANDLE " + _
                               "               FROM SAM_TIPOGUIA_MDGUIA MD " + _
						       "                    JOIN SAM_TIPOGUIA TG ON (TG.HANDLE = MD.TIPOGUIA) " + _
                               "              WHERE MD.REGIMEPAGTO = 'C' AND TG.TIPOGUIATISS = @TIPOPEGTISS) " + _
                               "  )"

  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  'SMS - 100620 - Gabriel - Tratamento do tipo de PEG para importação via Web Service e geração de PEG pelo autorizador externo.
  If ((CurrentQuery.FieldByName("ASSISTENCIA").AsString = "O") _
        And (CurrentQuery.FieldByName("TIPOPEGTISS").AsString <> "T") _
        And (CurrentQuery.FieldByName("TIPOPEGTISS").AsString <> "N")) _
      Or ((CurrentQuery.FieldByName("ASSISTENCIA").AsString = "M") And (CurrentQuery.FieldByName("TIPOPEGTISS").AsString = "T")) Then

		bsShowMessage("Tipo TISS inválido para a Assistência selecionada", "E")
		CanContinue =False
		Exit Sub
  End If

  If (TABREGIMEPGTO.PageIndex <> 0) And (CurrentQuery.FieldByName("TIPOPEGTISS").AsString <> "N") Then
		bsShowMessage("O regime de pagamento deve ser somente Credenciamento quando for selecionado um Tipo TISS diferente de 'Nenhum'", "E")
		CanContinue =False
		Exit Sub
  End If


  If (CurrentQuery.FieldByName("TIPOPEGTISS").AsString <> "N") And (CurrentQuery.FieldByName("MODELOGUIA").IsNull) Then
		bsShowMessage("Modelo de guia é obrigatório quando o Tipo TISS for diferente de 'Nenhum'", "E")
		CanContinue =False
		Exit Sub
  End If


  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT 1 FROM SAM_TIPOPEG WHERE TIPOPEGTISS = :TIPOPEG AND HANDLE <> :HANDLE")
  sql.ParamByName("TIPOPEG").AsString = CurrentQuery.FieldByName("TIPOPEGTISS").AsString
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If (Not sql.EOF) And (CurrentQuery.FieldByName("TIPOPEGTISS").AsString <> "N") Then
		bsShowMessage("Já existe um tipo de peg com o mesmo Tipo TISS selecionado", "E")
		CanContinue =False
		Exit Sub
  End If

 'SMS - 100620 - FIM


  Dim qAux As Object
  Set qAux = NewQuery

  qAux.Clear
  qAux.Add("SELECT HANDLE")
  qAux.Add("  FROM SAM_TIPOPEG")
  qAux.Add(" WHERE CODIGOEXPORTACAO = :CODIGOEXPORTACAO")
  qAux.Add("   AND HANDLE <> :HANDLE")

  qAux.ParamByName("CODIGOEXPORTACAO").AsString = CurrentQuery.FieldByName("CODIGOEXPORTACAO").AsString
  qAux.ParamByName("HANDLE").AsString = CurrentQuery.FieldByName("HANDLE").AsString

  qAux.Active = True

  If Not(qAux.FieldByName("HANDLE").IsNull) Then
    bsShowMessage("Já existe um Tipo de PEG com o código informado", "E")
    CODIGOEXPORTACAO.SetFocus
    CanContinue = False
  End If

End Sub

Public Sub TIPOPEGTISS_OnChange()
	CurrentQuery.FieldByName("MODELOGUIA").Value = Null
End Sub
