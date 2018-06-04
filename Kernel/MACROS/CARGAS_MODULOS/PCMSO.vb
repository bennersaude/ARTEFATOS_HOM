'HASH: 5C663DCC35A71AC80BA24D33D575218F
 

Public Sub GERAREXAMES_OnClick()
  Dim BSMed001 As Object
  Set BSMed001 = CreateBennerObject("BSMed001.GerarExames")
  BSMed001.Exec(CurrentSystem)

  Set BSMed001 = Nothing
End Sub
