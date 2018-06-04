'HASH: 5D6A4048E633FBC46828CDC68C51253D
'Macro: SFN_ROTINAFINFAT_INSS

Public Sub BOTAOCANCELAR_OnClick()
  Dim INSS As Object

  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If

  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT SITUACAO")
  SQL.Add("FROM SFN_ROTINAFIN")
  SQL.Add("WHERE HANDLE = :HROTINAFIN")
  SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString = "A" Then
    MsgBox("A Rotina não foi processada")
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing

  Set INSS = CreateBennerObject("SAMINSS.INSS")
  INSS.Cancelar(CurrentSystem)
  Set INSS = Nothing

  WriteAudit("C", HandleOfTable("SFN_ROTINAFINFAT_INSS"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Faturamento de INSS - Cancelamento")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim PRIMEIRODIA As Date
  Dim ULTIMODIA As Date
  Dim INSS As Object

  If CurrentQuery.State <>1 Then
    MsgBox("Os parâmetros não podem estar em edição")
    Exit Sub
  End If

  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT SITUACAO")
  SQL.Add("FROM SFN_ROTINAFIN")
  SQL.Add("WHERE HANDLE = :HROTINAFIN")
  SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString = "P" Then
    MsgBox("A Rotina já foi processada")
    Set SQL = Nothing
    Exit Sub
  End If

  Set INSS = CreateBennerObject("SAMINSS.INSS")
  INSS.Faturar(CurrentSystem)
  Set INSS = Nothing

  WriteAudit("P", HandleOfTable("SFN_ROTINAFINFAT_INSS"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Faturamento de INSS - Processamento")
End Sub

