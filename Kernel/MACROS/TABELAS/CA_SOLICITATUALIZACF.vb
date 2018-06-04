'HASH: 4713A9C9583AE282F5326A936A7C1EDC

'############### CENTRAL DE ATENDIMENTO #################

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  Select Case CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger
    Case 1
      SQL.Clear
      SQL.Add("SELECT NOME FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
      SQL.Active = True
      CurrentQuery.FieldByName("NOMERESPONSAVEL").AsString = SQL.FieldByName("NOME").AsString
    Case 2
      SQL.Clear
      SQL.Add("SELECT NOME FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
      SQL.Active = True
      CurrentQuery.FieldByName("NOMERESPONSAVEL").AsString = SQL.FieldByName("NOME").AsString
    Case 3
      SQL.Clear
      SQL.Add("SELECT NOME FROM SFN_PESSOA WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PESSOA").AsInteger
      SQL.Active = True
      CurrentQuery.FieldByName("NOMERESPONSAVEL").AsString = SQL.FieldByName("NOME").AsString
  End Select
  Set SQL = Nothing
End Sub

Public Sub TABLE_NewRecord()
  Dim vANO As String
  Dim SEQUENCIA As Long
  vANO = Format(ServerDate, "yyyy")
  NewCounter("CA_ATEND", CDate(vANO), 1, SEQUENCIA)
  CurrentQuery.FieldByName("ANO").Value = ("01/01/" + vANO)
  CurrentQuery.FieldByName("NUMERO").Value = SEQUENCIA
End Sub


'###############################################################

