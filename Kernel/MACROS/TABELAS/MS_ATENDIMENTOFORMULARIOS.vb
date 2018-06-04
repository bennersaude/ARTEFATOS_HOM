'HASH: F46D53EDFD3B7048A30DA027CC570D6F
'MACRO: MS_ATENDIMENTOFORMULARIOS

Public Sub PESQUISA_OnClick()
  Dim Obj As Object
  Dim vEmpresa, vFilial, vFormulario, vPac, vAtendForm As Integer

  Set Obj = CreateBennerObject("BSCLI004.Rotinas")

  vPac = RecordHandleOfTable("MS_PACIENTES")
  vFormulario = CurrentQuery.FieldByName("FORMULARIO").AsInteger
  vAtendForm = CurrentQuery.FieldByName("HANDLE").AsInteger
  vEmpresa = CurrentQuery.FieldByName("EMPRESA").AsInteger
  vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger

  Obj.PesquisaFormulario(CurrentSystem, 3, vPac, vFormulario, vEmpresa, vFilial, vAtendForm)

  Set Obj = Nothing
End Sub

