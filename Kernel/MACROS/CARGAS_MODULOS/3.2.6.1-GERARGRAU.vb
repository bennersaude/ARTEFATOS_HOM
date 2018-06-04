'HASH: 0839F2392CEF4730A677885EA62B0D40

Public Sub BOTAOGERARGRAU_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="DESCRICAO|GRAU"
  vCampos ="Descrição|Grau"

  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SAM_GRAU",vColunas,vCampos,"SAM_TIPOPRESTADOR_GRAU","TIPOPRESTADOR",RecordHandleOfTable("SAM_TIPOPRESTADOR"),"GRAU","Seleciona Grau")
  RefreshNodesWithTable("SAM_TIPOPRESTADOR_GRAU")
  Set interface =Nothing

End Sub
