'HASH: ADD1112FE7F74E6F2D5CCD56F04A3D9E


Public Sub CONSULTAR_OnClick()
  Dim Obj As Object
  Dim sql As Object
  Dim vEmpresa, vFilial, vFormulario, vPac, vAtendForm As Integer

  Set Obj = CreateBennerObject("BSCLI004.Rotinas")

  Set sql = NewQuery
  Set sql.Active = False
  sql.Add("SELECT FORMULARIO, HANDLE, EMPRESA, FILIAL FROM MS_ATENDIMENTOFORMULARIOS WHERE HANDLE = :FORMULARIO")
  sql.ParamByName("FORMULARIO").Value = RecordHandleOfTable("MS_ATENDIMENTOFORMULARIOS")
  sql.Active = True

  vPac = RecordHandleOfTable("MS_PACIENTES")
  vFormulario = sql.FieldByName("FORMULARIO").AsInteger
  vAtendForm = sql.FieldByName("HANDLE").AsInteger
  vEmpresa = sql.FieldByName("EMPRESA").AsInteger
  vFilial = sql.FieldByName("FILIAL").AsInteger

  Obj.PesquisaFormulario(CurrentSystem, 2, vPac, vFormulario, vEmpresa, vFilial, vAtendForm)

  Set Obj = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery
  sql.Clear
  sql.Add("DELETE FROM MS_ATENDIMENTORESPOSTAS WHERE FORMULARIO = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL
End Sub

