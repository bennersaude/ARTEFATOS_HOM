'HASH: EA4551A1A12F1FA8076CC22270FB850F
'#Uses "*bsShowMessage"

Option Explicit


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  If CurrentQuery.FieldByName("ESTADO").AsString <> "ZF" Then
    sql.Add("SELECT COUNT(1) QTD FROM ESTADOS WHERE SIGLA = :SIGLA")
    sql.ParamByName("SIGLA").AsString = CurrentQuery.FieldByName("ESTADO").AsString
    sql.Active = True

    If sql.FieldByName("QTD").AsInteger = 0 Then
      BsShowMessage("A sigla digitada é inválida!", "E")
      CanContinue = False
    End If
  End If

  sql.Clear
  sql.Add("SELECT COUNT(1) QTD FROM SAM_MATMED_ALIQUOTA WHERE ESTADO = :ESTADO AND GENERICO = :GENERICO AND HANDLE <> :HANDLE")
  sql.ParamByName("ESTADO").AsString = CurrentQuery.FieldByName("ESTADO").AsString
  sql.ParamByName("GENERICO").AsString = CurrentQuery.FieldByName("GENERICO").AsString
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If sql.FieldByName("QTD").AsInteger > 0 Then
    BsShowMessage("Já existe uma aliquota configurada para o estado e o tipo de medicamento selecionado!", "E")
    CanContinue = False
  End If

  Set sql = Nothing

End Sub
