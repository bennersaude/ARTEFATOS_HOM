'HASH: C2BC0DC4E597F8A0B4B9814254AFCC16
 
Public Sub BOTAOEXCLUIESP_OnClick()
  Dim Obj As Object

  Set Obj =CreateBennerObject("SamGerarCID.Especialidade")
  Obj.Excluir(CurrentSystem,CurrentQuery.FieldByName("ITEM").AsInteger)

  Set Obj =Nothing
End Sub

Public Sub BOTAOINCLUIESP_OnClick()
  Dim Obj As Object

  Set Obj =CreateBennerObject("SamGerarCID.Especialidade")
  Obj.Cadastrar(CurrentSystem,CurrentQuery.FieldByName("ITEM").AsInteger)

  Set Obj =Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim QTmp As Object
  Set QTmp = NewQuery

  QTmp.Active = False
  QTmp.Clear
  QTmp.Add("SELECT COUNT(*) QT FROM GER_ITEM_TIPOTRATAMENTO ")
  QTmp.Add(" WHERE ITEM = :ITEM ")
  QTmp.ParamByName("ITEM").AsInteger = CurrentQuery.FieldByName("ITEM").AsInteger
  QTmp.Active = True
  If QTmp.FieldByName("QT").AsInteger > 0 Then
    MsgBox("Não é possível incluir Especialidades! Já existem Tipo de Tratamento cadastradas para este item")
    CanContinue = False
  End If
  Set QTmp = Nothing





  Dim q As Object
  Set q = NewQuery
  q.Clear

  q.Add("SELECT 1 FROM ger_item_especialidade WHERE especialidade=:especialidade AND item=:item ")

  q.ParamByName("especialidade").Value = CurrentQuery.FieldByName("especialidade").Value
  q.ParamByName("item"         ).Value = CurrentQuery.FieldByName("item"         ).Value
  q.Active = True

  If Not q.EOF Then
    CanContinue = False
   	MsgBox("Especialidade já cadastrada para esse item !")
  End If
  Set q = Nothing






End Sub
