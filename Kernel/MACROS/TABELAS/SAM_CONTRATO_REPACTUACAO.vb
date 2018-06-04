﻿'HASH: CEB66053017550DE3DC836C6D5F6564D
'Macro: SAM_CONTRATO_REPACTUACAO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim SQL2 As Object

  Set SQL = NewQuery
  Set SQL2 = NewQuery

  SQL.Add("SELECT REPACTUACAO FROM SAM_CONTRATO")
  SQL.Add("WHERE HANDLE = :HCONTRATO")
  SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
  SQL.Active = True

  If SQL.FieldByName("REPACTUACAO").AsString = "N" Then
    bsShowMessage("Este Contrato não foi repactuado. Inclusão não permitida", "E")
    CanContinue = False
    Set SQL = Nothing
    Set SQL2 = Nothing
    Exit Sub
  End If

  SQL.Clear
  SQL.Add("SELECT CONTRATOMODADESAOPRC, IDADEMAXIMA, CONTRATOTPDEP")
  SQL.Add("FROM SAM_CONTRATO_MODADESAOPRC_FX")
  SQL.Add("WHERE HANDLE = :HMODADESAOPRCFX")
  SQL.ParamByName("HMODADESAOPRCFX").Value = CurrentQuery.FieldByName("CONTRATOMODADESAOPRCFX").AsInteger
  SQL.Active = True

  SQL2.Add("SELECT HANDLE, IDADEMAXIMA")
  SQL2.Add("FROM SAM_CONTRATO_MODADESAOPRC_FX")
  SQL2.Add("WHERE CONTRATOMODADESAOPRC = :HMODADESAOPRC")
  SQL2.Add("  AND IDADEMAXIMA >= 59")
  SQL2.Add("  AND IDADEMAXIMA < :IDADE")

  If Not SQL.FieldByName("CONTRATOTPDEP").IsNull Then
    SQL2.Add("  AND CONTRATOTPDEP = :HCONTRATOTPDEP")
    SQL2.Add("ORDER BY IDADEMAXIMA DESC")
    SQL2.ParamByName("HCONTRATOTPDEP").Value = SQL.FieldByName("CONTRATOTPDEP").AsInteger
  Else
    SQL2.Add("ORDER BY IDADEMAXIMA DESC")
  End If

  SQL2.ParamByName("HMODADESAOPRC").Value = SQL.FieldByName("CONTRATOMODADESAOPRC").AsInteger
  SQL2.ParamByName("IDADE").Value = SQL.FieldByName("IDADEMAXIMA").AsInteger

  SQL2.Active = True

  If SQL2.EOF Then
    bsShowMessage("Não é permitido incluir repactuação para esta faixa", "E")
    CanContinue = False
    Set SQL = Nothing
    Set SQL2 = Nothing
    Exit Sub
  End If

  If CurrentQuery.FieldByName("IDADE").AsInteger <= SQL2.FieldByName("IDADEMAXIMA").AsInteger Then
    bsShowMessage("A idade deve ser maior que a idade máxima da faixa anterior", "E")
    CanContinue = False
    Set SQL = Nothing
    Set SQL2 = Nothing
    Exit Sub
  End If

  If CurrentQuery.FieldByName("IDADE").AsInteger > SQL.FieldByName("IDADEMAXIMA").AsInteger Then
    bsShowMessage("A Idade não pode ser maior que a idade máxima da faixa", "E")
    CanContinue = False
    Set SQL = Nothing
    Set SQL2 = Nothing
    Exit Sub
  End If

  Set SQL = Nothing
  Set SQL2 = Nothing
End Sub

