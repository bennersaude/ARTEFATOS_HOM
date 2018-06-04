'HASH: 4F14E0407A7D18676A1B1623CFA6B427

'##################$ CENTRAL DE ATENDIMENTO ####################


Public Sub TABLE_NewRecord()
  Dim ANO As Date
  Dim SEQUENCIA As Long
  NewCounter("CA_ATEND", ANO, 1, SEQUENCIA)
  CurrentQuery.FieldByName("ANO").Value = DatePart("yyyy", ANO)
  CurrentQuery.FieldByName("NUMERO").Value = SEQUENCIA
End Sub


'###############################################################


