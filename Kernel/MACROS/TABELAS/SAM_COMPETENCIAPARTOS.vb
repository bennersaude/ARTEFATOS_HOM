'HASH: 87EAF676662D7B08E585D14EBD76244D
'Macro: SAM_COMPETENCIAPARTOS
'#Uses "*bsShowMessage"

Public Sub BOTAOPROCESSAR_OnClick()
 'SMS 811825 - RN 368
 Dim var_ano As Integer 'Competencia
 Dim var_mes As Integer

 var_ano=Year(ServerNow)
 var_ano=var_ano-1

 If CurrentQuery.FieldByName("COMPETENCIA").AsString <> "" Then
    var_ano=CurrentQuery.FieldByName("COMPETENCIA").AsInteger
 End If

 var_mes=Month(ServerNow)

 Dim pCompetencia As String
 pCompetencia=Str(var_ano)

 Dim component As CSBusinessComponent

 Set component = BusinessComponent.CreateInstance("Benner.Saude.Adm.Processos.ProcessoContabilizaPartos, Benner.Saude.Adm.Processos")
 component.AddParameter(pdtString, pCompetencia)

  component.Execute("ContabilizarPartosAno")


 If CurrentQuery.FieldByName("COMPETENCIA").AsString = "" Then
  If var_mes <= 3  Then
    var_ano=var_ano-1
    pCompetencia=Str(var_ano)
    component.ClearParameters
    component.AddParameter(pdtString, pCompetencia)
    component.Execute("ContabilizarPartosAno")
  End If
 End If

Set component = Nothing

 bsShowMessage("Processado com sucesso!", "I")
End Sub



Public Sub TABLE_BeforePost(CanContinue As Boolean)


  Dim query As BPesquisa
  Set query = NewQuery
  query.Active = False
  query.Clear
  query.Add("SELECT HANDLE                      ")
  query.Add("  FROM SAM_COMPETENCIAPARTOS       ")
  query.Add("  WHERE COMPETENCIA=:PCOMPETENCIA  ")
  query.Add("    AND HANDLE <>:PHANDLE          ")
  query.ParamByName("PCOMPETENCIA").AsInteger= CurrentQuery.FieldByName("COMPETENCIA").AsInteger
  query.ParamByName("PHANDLE").AsInteger= CurrentQuery.FieldByName("HANDLE").AsInteger
  query.Active = True

  If Not query.EOF Then
   bsShowMessage("Competência já informada!", "I")
   CanContinue=False
  End If
  Set  query=Nothing
End Sub

Public Sub TABLE_AfterInsert()
 CurrentQuery.FieldByName("DATAHORAINCLUSAO").AsDateTime=ServerNow
 CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger=CurrentUser
End Sub
