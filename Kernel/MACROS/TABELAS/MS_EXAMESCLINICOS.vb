'HASH: C633B45ADA5F56D691EAF3E38586E4DF

Dim Empresa, FILIAL, Pac, EXAME, Handle As Integer
Dim DATA As Date



Public Sub HEMOGRAMA_OnClick()
  Set Obj = CreateBennerObject("MS.Hemograma")
  Empresa = CurrentQuery.FieldByName("EMPRESA").AsInteger
  FILIAL = CurrentQuery.FieldByName("FILIAL").AsInteger
  EXAME = CurrentQuery.FieldByName("HANDLE").AsInteger
  Pac = RecordHandleOfTable("MS_PACIENTES")
  Handle = RecordHandleOfTable("MS_EXAMESCLINICOS")
  DATA = CurrentQuery.FieldByName("DATA").AsDateTime
  Obj.Exec(CurrentSystem, Pac, Empresa, FILIAL, EXAME, Handle, DATA)

  Set obj = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Set Sql = NewQuery
  If CurrentQuery.FieldByName("EXAME").AsString <>"" Then
    Sql.Add "SELECT DESCRICAO FROM SAM_TGE WHERE HANDLE = " + CurrentQuery.FieldByName("EXAME").AsString
    Sql.Active = True
    If Not Sql.EOF Then

      If Left(Sql.FieldByName("DESCRICAO").AsString, 9) = "Hemograma" Then
        HEMOGRAMA.Visible = True
      End If
    End If
  End If
  Set Sql = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim Sql
  Set Sql = NewQuery
  If CurrentQuery.FieldByName("EXAME").AsString <>"" Then
    Sql.Add "SELECT DESCRICAO FROM SAM_TGE WHERE HANDLE = " + CurrentQuery.FieldByName("EXAME").AsString
    Sql.Active = True
    If Not Sql.EOF Then

      If Left(Sql.FieldByName("DESCRICAO").AsString, 9) = "Hemograma" Then
        HEMOGRAMA.Visible = True
      Else
        HEMOGRAMA.Visible = False
      End If
    End If
  End If
  Set Sql = Nothing
End Sub


Public Sub TABLE_NewRecord()
  If VisibleMode Then
    CurrentQuery.FieldByName("FATOGERADOR").Value = 6
    HEMOGRAMA.Visible = False
  End If
End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  SQL.Active = True

  If((SQL.FieldByName("DATAINICIAL").IsNull)Or((Not SQL.FieldByName("DATAINICIAL").IsNull)And(Not SQL.FieldByName("DATAFINAL").IsNull)))Then
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

  If((SQL.FieldByName("DATAINICIAL").IsNull)Or((Not SQL.FieldByName("DATAINICIAL").IsNull)And(Not SQL.FieldByName("DATAFINAL").IsNull)))Then
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

  If((SQL.FieldByName("DATAINICIAL").IsNull)Or((Not SQL.FieldByName("DATAINICIAL").IsNull)And(Not SQL.FieldByName("DATAFINAL").IsNull)))Then
  MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
  CanContinue = False
  Exit Sub
End If
End Sub

