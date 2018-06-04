'HASH: 73F780A5A4C779D2377AC88E504C44A0


Public Sub TABLE_AfterScroll()
  BAIRRO.ReadOnly = True
  BENEFICIARIO.ReadOnly = True
  CEP.ReadOnly = True
  COMPLEMENTO.ReadOnly = True
  CONTRATO.ReadOnly = True
  CORRESPONDENCIA.ReadOnly = True
  CORRESPONDENCIANUMERO.ReadOnly = True
  ESTADO.ReadOnly = True
  LOGRADOURO.ReadOnly = True
  MUNICIPIO.ReadOnly = True
  NUMERO.ReadOnly = True
  PESSOA.ReadOnly = True
  ROTINACONTRATO.ReadOnly = True
  TABTIPORESPONSAVEL.ReadOnly = True
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim APAGA As Object
  Set APAGA = NewQuery

  APAGA.Add("DELETE FROM SAM_ROTAVISOSUSPENSAO_ATRASO")
  APAGA.Add(" WHERE ROTINADEST = :HANDLEDEST")
  APAGA.ParamByName("HANDLEDEST").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  APAGA.ExecSQL

  Set APAGA = Nothing
End Sub

