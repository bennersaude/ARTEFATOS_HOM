'HASH: 2FC7550237B6BE1E404F747D3B24809C
 

Public Sub BOTAOGERAREGIMEATEND_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="CODIGO|DESCRICAO"

  vCampos ="Código|Descrição"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_REGIMEATENDIMENTO",vColunas,vCampos,"SFN_REGRAPAG_REGIMEATENDIMENTO","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"REGIMEATENDIMENTO","Seleciona Regime de Atendimento")
  Set interface =Nothing
End Sub
