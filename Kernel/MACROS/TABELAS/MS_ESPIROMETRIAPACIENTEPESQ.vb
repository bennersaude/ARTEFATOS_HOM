'HASH: C9498A4A563DFDFA6A00C07102EEBAF0

Dim Obj, SqlForm, Sql2 As Object
Dim Empresa, FILIAL, Funcionario, Pac, Tipo As Integer

Public Sub CONSULTAR_OnClick()
  Set Obj = CreateBennerObject("BSMed001.RealizarPesquisa")
  Set Sql2 = NewQuery
  Empresa = RecordHandleOfTable("EMPRESA")
  FILIAL = RecordHandleOfTable("FILIAL")

  Pac = RecordHandleOfTable("MS_PACIENTES")
  If Pac <= 0 Then
    Sql2.Clear
    Sql2.Add("SELECT PACIENTE FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
    Sql2.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
    Sql2.Active = True
    Pac = Sql2.FieldByName("PACIENTE").AsInteger
  End If
  Tipo = 2
  Exame = 4

  Sql = "SELECT FORMULARIO FROM MS_ESPIROMETRIAFORMULARIOS WHERE HANDLE = :FORMULARIO"
  Set SqlForm = NewQuery
  SqlForm.Add(Sql)
  SqlForm.ParamByName("FORMULARIO").Value = RecordHandleOfTable("MS_ESPIROMETRIAFORMULARIOS")
  SqlForm.Active = True

  Obj.Data(CurrentSystem)CurrentQuery.FieldByName("HANDLE").AsInteger
  Obj.Formulario(CurrentSystem)SqlForm.FieldByName("FORMULARIO").AsInteger
  Obj.Exec(CurrentSystem, Pac, Empresa, FILIAL, Tipo, Exame)

  Set Obj = Nothing
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("DELETE FROM MS_ESPIROMETRIAPACPESQRESP WHERE FORMULARIO = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL
End Sub

