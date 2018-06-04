'HASH: F5E314B9859BD96AA1984C52485CCE06


Public Sub AFASTAR_OnClick()
  'Set obj = CreateBennerObject("RH.Afastar")
  'obj.Dataafastamento(CurrentSystem)CurrentQuery.FieldByName("DATAAFASTAMENTO").AsDateTime
  If CurrentQuery.FieldByName("DATAAFASTAMENTO").IsNull Then
    MsgBox "Paciente não pode ser afastado, está sem data de afastamento no atestado"
  'Else
  '  obj.Dataretorno(CurrentSystem)CurrentQuery.FieldByName("DATAAFASTAMENTO").AsDateTime + CurrentQuery.FieldByName("AFASTADO").AsInteger
  '  obj.Diasatestado(CurrentSystem)CurrentQuery.FieldByName("AFASTADO").AsInteger
  '  obj.Paciente(CurrentSystem)RecordHandleOfTable("MS_PACIENTES")
  '  obj.Codigo(CurrentSystem)1
  '  obj.Exec(CurrentSystem)
  End If
  'Set obj = Nothing

End Sub

Public Sub CIDS_OnClick()
  Set obj = CreateBennerObject("BSMed001.CidsAtestado")
  obj.Exec(CurrentSystem)
  Set obj = Nothing
End Sub

Public Sub EXCLUIR_OnClick()
  Dim obj As Object
  Set obj = CreateBennerObject("BSMed001.ExcluirAtestado")
  obj.Exec(CurrentSystem)
  Set obj = Nothing
End Sub

Public Sub PACIENTE_OnExit()
End Sub

Public Sub RETORNAR_OnClick()
  'Dim obj As Object
  'Set obj = CreateBennerObject("RH.RetornarAfastamento")
  'obj.Paciente(CurrentSystem)RecordHandleOfTable("MS_PACIENTES")
  'obj.Codigo(CurrentSystem)1

'  obj.Exec(CurrentSystem)
 ' Set obj = Nothing
End Sub

Public Sub TABLE_AfterPost()
  '  CIDS.Visible =True
  EXCLUIR.Visible = True
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("MEDICO").IsNull Then
    MsgBox "Médico responsável deve ser indicado!"
    CanContinue = False
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

'  CIDS.Visible =False
EXCLUIR.Visible = False
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Clear
  Sql.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  Sql.Active = True

  If((Sql.FieldByName("DATAINICIAL").IsNull)Or((Not Sql.FieldByName("DATAINICIAL").IsNull)And(Not Sql.FieldByName("DATAFINAL").IsNull)))Then
  MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
  CanContinue = False
  Exit Sub
End If

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Clear
  Sql.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  Sql.Active = True

  If((Sql.FieldByName("DATAINICIAL").IsNull)Or((Not Sql.FieldByName("DATAINICIAL").IsNull)And(Not Sql.FieldByName("DATAFINAL").IsNull)))Then
  MsgBox("Só é possível alterar um atendimento que esteja em aberto!")
  CanContinue = False
  Exit Sub
End If
End Sub


Public Sub TABLE_AfterScroll()
  Dim Sql As Object
  Set Sql = NewQuery

  EXCLUIR.Enabled = True

  Sql.Clear
  Sql.Add("SELECT DATAFINAL, DATAINICIAL FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
  Sql.Active = True

  If((Sql.FieldByName("DATAINICIAL").IsNull)Or((Not Sql.FieldByName("DATAINICIAL").IsNull)And(Not Sql.FieldByName("DATAFINAL").IsNull)))Then
  EXCLUIR.Enabled = False
  Exit Sub
End If

End Sub

