'HASH: 4E214A7EB8BDACA2083762CBD68E8CFA



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
  Exame = 3
  Obj.Exec(CurrentSystem, Pac, Empresa, FILIAL, Tipo, Exame)

  Set Obj = Nothing
End Sub

'Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
'Dim Sql As Object
'Set Sql =NewQuery
'  Sql.Clear
'  Sql.Add("SELECT 1 FROM MS_AUDIOMETRIAPACPESQRESP WHERE HANDLE = :HANDLE")
'  Sql.ParamByName("HANDLE").AsInteger =CurrentQuery.FieldByName("HANDLE").AsInteger
'  Sql.Active =True
'  If Not Sql.EOF Then
'    MsgBox("Existem respostas para este formulário!")
'    CanContinue =False
'    Exit Sub
'  End If

'  Sql.Clear
'  Sql.Add("DELETE FROM MS_AUDIOMETRIAPACIENTEPESQUISA WHERE FORMULARIO = :HANDLE")
'  Sql.ParamByName("HANDLE").AsInteger =CurrentQuery.FieldByName("HANDLE").AsInteger
'  Sql.ExecSQL

'End Sub
