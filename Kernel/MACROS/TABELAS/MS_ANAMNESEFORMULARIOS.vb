'HASH: 53AE0BA2BE15007D23AE528E82DBA940

Dim Obj, Sql As Object
Dim Empresa, Filial, Funcionario, Pac, Tipo As Integer



Public Sub PESQUISA_OnClick()
  Set Obj = CreateBennerObject("BSMed001.RealizarPesquisa")
  Set Sql = NewQuery
  Empresa = CurrentQuery.FieldByName("EMPRESA").AsInteger
  Filial = CurrentQuery.FieldByName("FILIAL").AsInteger

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
  Exame = 1
  Obj.Exec(CurrentSystem, Pac, Empresa, Filial, Tipo, Exame)

  Set Obj = Nothing
End Sub

'Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
'Dim Sql As Object
'Set Sql =NewQuery
'  Sql.Clear
'  Sql.Add("SELECT 1 FROM MS_ANAMNESEPACIENTEPESQUISARES WHERE HANDLE = :HANDLE")
'  Sql.ParamByName("HANDLE").AsInteger =CurrentQuery.FieldByName("HANDLE").AsInteger
'  Sql.Active =True
'  If Not Sql.EOF Then
'    MsgBox("Existem respostas para este formulário!")
'    CanContinue =False
'    Exit Sub
'  End If

'  Sql.Clear
'  Sql.Add("DELETE FROM MS_ANAMNESEPACIENTEPESQUISA WHERE FORMULARIO = :HANDLE")
'  Sql.ParamByName("HANDLE").AsInteger =CurrentQuery.FieldByName("HANDLE").AsInteger
'  Sql.ExecSQL
'End Sub
