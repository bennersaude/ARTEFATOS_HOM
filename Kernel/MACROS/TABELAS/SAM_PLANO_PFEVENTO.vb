'HASH: 56089621F860428B7B45AAE40A603F0D
'Macro: SAM_PLANO_PFEVENTO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'SMS 40817 - Anderson Lonardoni - 03/05/2005
  'SE O TIPO DE CONTAGEM FOR NO CONTRATO O TIPO DE PERÍODO DEVERÁ SER CIVIL OU POR ADESÃO DO CONTRATO
  If (CurrentQuery.FieldByName("TIPOCONTAGEM").AsString = "C") _
       And ((CurrentQuery.FieldByName("TIPOPERIODO").AsString = "F") Or (CurrentQuery.FieldByName("TIPOPERIODO").AsString = "B")) _
       And (CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 2) Then
    CanContinue = False
    bsShowMessage("Tipo de contagem no contrato exige que o tipo de período seja civil ou por adesão do contrato!", "E")
  End If
  'SE O TIPO DE CONTAGEM FOR NA FAMÍLIA O TIPO DE PERÍODO DEVERÁ SER CIVIL, POR ADESÃO DO CONTRATO OU DA FAMÍLIA
  If (CurrentQuery.FieldByName("TIPOCONTAGEM").AsString = "F") And (CurrentQuery.FieldByName("TIPOPERIODO").AsString = "B") Then
    CanContinue = False
    bsShowMessage("Tipo de contagem na família exige que o tipo de período seja civil, por adesão do contrato ou da família!", "E")
  End If
End Sub



