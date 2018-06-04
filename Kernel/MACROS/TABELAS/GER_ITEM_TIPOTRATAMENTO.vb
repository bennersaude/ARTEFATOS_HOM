'HASH: AF499BD6B9AE883007D98241700FAD14
 
Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim QTmp As Object
  Set QTmp = NewQuery

  QTmp.Active = False
  QTmp.Clear
  QTmp.Add("SELECT COUNT(*) QT FROM GER_ITEM_ESPECIALIDADE ")
  QTmp.Add(" WHERE ITEM = :ITEM ")
  QTmp.ParamByName("ITEM").AsInteger = CurrentQuery.FieldByName("ITEM").AsInteger
  QTmp.Active = True
  If QTmp.FieldByName("QT").AsInteger > 0 Then
    MsgBox("Não é possível incluir Tipo de Tratamento! Já existem Especialidades cadastradas para este item")
    CanContinue = False
  End If
  Set QTmp = Nothing



  Dim q As Object
  Set q = NewQuery

  q.Active = False
  q.Clear
  q.Add("select 1 from ger_item_tipotratamento where tipotratamento=:tipotratamento and item=:item ")
  q.ParamByName("tipotratamento").Value = CurrentQuery.FieldByName("tipotratamento").Value
  q.ParamByName("item"          ).Value = CurrentQuery.FieldByName("item"          ).Value

  q.Active = True

  If Not q.EOF Then
    CanContinue = False
   	MsgBox("Este Tipo de Tratamento já está cadastrado para este Item !")
  End If
  Set q = Nothing



End Sub
