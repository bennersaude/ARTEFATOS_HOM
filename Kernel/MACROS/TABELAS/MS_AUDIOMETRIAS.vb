'HASH: 7B7363380E7D2E70573B42479D3D455E

Public Sub PARAMETRIZARGRAFICO_OnClick()
  Set obj = CreateBennerObject("BSMed001.PersonalizarConfig")
  obj.Exec(CurrentSystem, CurrentQuery.FieldByName("MATRICULA").AsInteger)
  Set obj = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Num As Integer
  Dim Sql

  If VisibleMode Then
    Num = CurrentQuery.FieldByName("HZ500").Value
    CanContinue = ConfereValor(Num)
    CanContinue = True
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ1000").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ2000").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ3000").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ4000").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ6000").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ8000").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ500E").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ1000E").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ2000E").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ3000E").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ4000E").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ6000E").Value
      CanContinue = ConfereValor(Num)
    End If
    If CanContinue = True Then
      Num = CurrentQuery.FieldByName("HZ8000E").Value
      CanContinue = ConfereValor(Num)
    End If


    If CurrentQuery.FieldByName("PADRAO").AsBoolean = True Then
      Set Sql = NewQuery
      Sql.Add "SELECT PADRAO FROM MS_AUDIOMETRIAS WHERE EMPRESA = " + CurrentQuery.FieldByName("EMPRESA").AsString + " AND FILIAL = " + CurrentQuery.FieldByName("FILIAL").AsString + " AND PACIENTE = " + CurrentQuery.FieldByName("PACIENTE").AsString + " AND PADRAO = 'S'
      Sql.Active = True
      If Not Sql.EOF Then
        MsgBox "Paciente já possui exame de referência!"
        CanContinue = False
      End If
      If CurrentQuery.FieldByName("DATAINICIO").IsNull Then
        MsgBox "Data de início do exame de referência deve ser informada!"
        CanContinue = False
      End If
      Set Sql = Nothing
    Else
      If (Not CurrentQuery.FieldByName("DATAINICIO").IsNull) Then
        MsgBox "Data de início do exame de referência inválida - pois esse exame não é de referência!"
        CanContinue = False
      End If

    End If
  End If
End Sub

Function ConfereValor(Num As Integer)As Boolean
  If CurrentQuery.FieldByName("UTILIZA").Value = 1 Then
    If Num <> (0) And Num <> (10) And Num <> (20) And Num <> (25) And Num <> (30) And Num <> (40) And Num <> (50) And Num <> (55) And Num <> (60) And Num <> (70) And Num <> (80) And Num <> (90) And Num <> (100) And Num <> (110) Then
      MsgBox "Verifique valor informado, não válido para o exame!"
      ConfereValor = False
    Else
      ConfereValor = True
    End If
  Else
    If Num <> ( -10) And Num <> ( -5) And Num <> (0) And Num <> (5) And Num <> (10) And Num <> (15) And Num <> (20) And Num <> (25) And Num <> (30) And Num <> (35) And Num <> (40) And Num <> (45) And Num <> (50) And Num <> (55) And Num <> (60) And Num <> (65) And Num <> (70) And Num <> (75) And Num <> (80) And Num <> (85) And Num <> (90) And Num <> (95) And Num <> (100) And Num <> (105) And Num <> (110) And Num <> (115) And Num <> (120) Then
      MsgBox "Verifique valor informado, não válido para o exame!"
      ConfereValor = False
    Else
      ConfereValor = True
    End If
  End If
End Function



Public Sub VISUALIZARGRAFICO_OnClick()
  If CurrentQuery.FieldByName("UTILIZA").Value = 1 Then
    Set obj = CreateBennerObject("BSMed001.AudiometriaGraficos")
  Else
    Set obj = CreateBennerObject("BSMed001.AudiometriaGraficosInss")
  End If
  obj.Exec(CurrentSystem, CurrentQuery.FieldByName("MATRICULA").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set obj = Nothing
End Sub


Public Sub TABLE_NewRecord()
  If VisibleMode Then
    CurrentQuery.FieldByName("FATOGERADOR").Value = 6
  End If
End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If ((SQL.FieldByName("DATAINICIAL").IsNull) Or ((Not SQL.FieldByName("DATAINICIAL").IsNull) And (Not SQL.FieldByName("DATAFINAL").IsNull))) Then
    MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
    CanContinue = False
    Exit Sub
  End If
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If ((SQL.FieldByName("DATAINICIAL").IsNull) Or ((Not SQL.FieldByName("DATAINICIAL").IsNull) And (Not SQL.FieldByName("DATAFINAL").IsNull))) Then
    MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
    CanContinue = False
    Exit Sub
  End If

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If ((SQL.FieldByName("DATAINICIAL").IsNull) Or ((Not SQL.FieldByName("DATAINICIAL").IsNull) And (Not SQL.FieldByName("DATAFINAL").IsNull))) Then
    MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
    CanContinue = False
    Exit Sub
  End If
End Sub

