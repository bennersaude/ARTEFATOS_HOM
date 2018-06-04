'HASH: 999DAC906D495C61CDDF3146DAF5F209
 

Public Sub GERACLASSE_OnClick()

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  vColunas ="ESTRUTURA|DESCRICAO"

  vCampos ="Estrutura|Desrição"
'Parâmetros(pTabela                               ,pColuna,pCampos,pTabAssoc                   ,pCampoAssoc    ,pHandleAssoc                             ,pCampoAssoc2   ,pTitulo: WideString),pSqlEspecial    ,pMostrar
      'InterfacePrestador.sELECIONA("SAM_PRESTADOR"     ,vColunas,vCampos,"SAM_CONTRATO_PRECO_PRESTADOR","CONTRATOPRECO" ,RecordHandleOfTable("SAM_CONTRATO_PRECO"),"PRESTADOR"    ,"Lista de Prestadores",vMenosPrestadores,vMostrar)
      'interface.sELECIONA        ("SFN_CLASSEGERENCIAL",vColunas,vCampos,"SFN_ITEMNOTA_CLASSEGERENCIAL","ITEMNOTA"      ,RecordHandleOfTable("SFN_ITEMNOTA")     ,"CLASSEGERENCIAL","Seleciona Classe gerencial","","S")
  Set interface =CreateBennerObject("Procura.Procurar")
  interface.sELECIONA(CurrentSystem,"SFN_CLASSEGERENCIAL",vColunas,vCampos,"SFN_ITEMNOTA_CLASSEGERENCIAL","ITEMNOTA",RecordHandleOfTable("SFN_ITEMNOTA"),"CLASSEGERENCIAL","Seleciona Classe gerencial","","S")
  Set interface =Nothing


End Sub
