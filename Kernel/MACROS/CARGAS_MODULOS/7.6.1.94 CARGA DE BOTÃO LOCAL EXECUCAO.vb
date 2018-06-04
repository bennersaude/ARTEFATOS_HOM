'HASH: 2208D3B4641EB649D5DD71554D27334A
 

Public Sub BOTAOGERALOCALEXECUC_OnClick()
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO"

  vCampos ="Descrição"

  Set Interface =CreateBennerObject("Procura.Procurar")
  Interface.SELECIONA(CurrentSystem,"SAM_TIPOPRESTADOR",vColunas,vCampos,"SFN_REGRAPAG_TIPOPRESLOCALEXEC","REGRAPAG",RecordHandleOfTable("SFN_REGRAPAG"),"TIPOPRESTADOR","Seleciona Tipo de Prestador do Local de Execução")
  Set Interface =Nothing
End Sub
