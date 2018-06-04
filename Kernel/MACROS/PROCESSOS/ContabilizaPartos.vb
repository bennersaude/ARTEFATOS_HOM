'HASH: 4A2AF29D2E939B41FE1161BA50AD6E8C
'#Uses "*bsShowMessage"

Public Sub Main
 'ContabilizaPartos
 'SMS 811825 - RN 368
 Dim var_ano As Integer 'Competencia
 Dim var_mes As Integer

 var_ano=Year(ServerNow)
 var_mes=Month(ServerNow)

 var_ano=var_ano-1
 Dim pCompetencia As String
 pCompetencia=Str(var_ano)

 Dim component As CSBusinessComponent
 Dim handleProcesso As Long

 Set component = BusinessComponent.CreateInstance("Benner.Saude.Adm.Processos.ProcessoContabilizaPartos, Benner.Saude.Adm.Processos")
 component.AddParameter(pdtString, pCompetencia)

 component.Execute("ContabilizarPartosAno")

 If var_mes <= 3  Then
    var_ano=var_ano-1
    pCompetencia=Str(var_ano)
    component.ClearParameters
    component.AddParameter(pdtString, pCompetencia)
    component.Execute("ContabilizarPartosAno")
 End If

 Set component = Nothing

End Sub
