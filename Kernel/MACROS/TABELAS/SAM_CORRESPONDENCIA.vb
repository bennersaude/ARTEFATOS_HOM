'HASH: 77F9560E5450DB2D0910A4282D2FEFA2
'macro SAM_CORRESPONDENCIA

Public Sub TABLE_AfterScroll()
  BENEFICIARIO.ReadOnly = True
  ANO.ReadOnly = True
  CONTRATO.ReadOnly = True
  EMISSAODATA.ReadOnly = True
  EMISSAOUSUARIO.ReadOnly = True
  NUMERO.ReadOnly = True
  ORIGEM.ReadOnly = True
  PESSOA.ReadOnly = True
  PRESTADOR.ReadOnly = True
  RELATORIO.ReadOnly = True
  ROTINACORRESP.ReadOnly = True

  DEVOLUCAODATA.ReadOnly = False
  DEVOLUCAOMOTIVO.ReadOnly = False
End Sub

