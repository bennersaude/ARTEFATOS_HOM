'HASH: DC6FEDEC41A6102D52BFC959A03C1A45
 

Public Sub BOTAOGERAGRAU_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO|GRAU"

  vCampos ="Descrição|Grau"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_GRAU",vColunas,vCampos,"SFN_REGRAPAG_GRAU","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"GRAU","Seleciona Grau")
  Set interface =Nothing
End Sub
