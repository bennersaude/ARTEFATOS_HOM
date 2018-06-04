'HASH: 06C4C5F42E94361DF70727CC9C5FF8DD
 

Public Sub BOTAOGERAFINALIDADE_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO"

  vCampos ="Descrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_FINALIDADEATENDIMENTO",vColunas,vCampos,"SFN_REGRAPAG_FINALIDADEATEND","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"FINALIDADEATENDIMENTO","Seleciona Finalidade de Atendimento")
  Set interface =Nothing
End Sub
