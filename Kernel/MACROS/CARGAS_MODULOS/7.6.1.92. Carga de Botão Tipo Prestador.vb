'HASH: 594F17F0B4577EAA9E76214F5E6BFECB
 

Public Sub BOTAOGERATIPOPRESTAD_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO"

  vCampos ="Descrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_TIPOPRESTADOR",vColunas,vCampos,"SFN_REGRAPAG_TIPOPRESTADOR","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"TIPOPRESTADOR","Seleciona Tipo de Prestador")
  Set interface =Nothing
End Sub
