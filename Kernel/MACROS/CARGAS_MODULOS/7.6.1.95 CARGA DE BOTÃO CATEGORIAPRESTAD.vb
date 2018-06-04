'HASH: 2182F9C27639016B655D0613660535FA



Public Sub BOTAOCATEGORIAPRESTA_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO"

  vCampos ="Descrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_CATEGORIA_PRESTADOR",vColunas,vCampos,"SFN_REGRAPAG_CATEGORIAPREST","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"CATEGORIAPRESTADOR","Seleciona Categoria de Prestador")
  Set interface =Nothing
End Sub
