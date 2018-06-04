'HASH: 7E29804F25822F4DE4A2E4CB99DE6D8E

Dim Obj, Sql As Object
Dim Empresa, FILIAL, Funcionario, Pac, Tipo As Integer



Public Sub PESQUISA_OnClick()
  Set Obj = CreateBennerObject("BSMed001.RealizarPesquisa")
  Set Sql = NewQuery
  Empresa = CurrentQuery.FieldByName("EMPRESA").AsInteger
  FILIAL = CurrentQuery.FieldByName("FILIAL").AsInteger

  Obj.Formulario(CurrentSystem)CurrentQuery.FieldByName("FORMULARIO").AsInteger
  Obj.Data(CurrentSystem)1
  Pac = RecordHandleOfTable("MS_PACIENTES")
  If Pac <= 0 Then
    Sql.Clear
    Sql.Add("SELECT PACIENTE FROM MS_ATENDIMENTOS WHERE HANDLE = :HANDLE")
    Sql.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("MS_ATENDIMENTOS")
    Sql.Active = True
    Pac = Sql.FieldByName("PACIENTE").AsInteger
  End If
  Tipo = 0
  Exame = 4
  Obj.Exec(CurrentSystem, Pac, Empresa, FILIAL, Tipo, Exame)

  Set Obj = Nothing
End Sub

'Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
'Dim Sql As Object
'Set Sql =NewQuery
'  Sql.Clear
'  Sql.Add("SELECT 1 FROM MS_ESPIROMETRIAPACPESQRESP WHERE HANDLE = :HANDLE")
'  Sql.ParamByName("HANDLE").AsInteger =CurrentQuery.FieldByName("HANDLE").AsInteger
'  Sql.Active =True
'  If Not Sql.EOF Then
'    MsgBox("Existem respostas para este formulário!")
'    CanContinue =False
'    Exit Sub
'  End If

'  Sql.Clear
'  Sql.Add("DELETE FROM MS_ESPIROMETRIAPACIENTEPESQ WHERE FORMULARIO = :HANDLE")
'  Sql.ParamByName("HANDLE").AsInteger =CurrentQuery.FieldByName("HANDLE").AsInteger
'  Sql.ExecSQL

'End Sub
