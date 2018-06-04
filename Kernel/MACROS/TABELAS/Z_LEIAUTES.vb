'HASH: 9C0F1D2459A85B2FDBE4CB5EC1DFF5E1
'#Uses "*bsShowMessage"

Public Sub TABELA_OnExit()
  Dim QWork As Object, sDesc
  If CurrentQuery.FieldByName("DESCRICAO").IsNull Then
    Set QWork = NewQuery
    If Not CurrentQuery.FieldByName("LEIAUTE").IsNull Then
      QWork.Add("SELECT A.DESCRICAO PAI, B.NOME TABELA FROM Z_LEIAUTES A, Z_TABELAS B WHERE A.HANDLE = :LAYPAI AND B.HANDLE = :LAY")
      QWork.ParamByName("LAYPAI").Value = CurrentQuery.FieldByName("LEIAUTE").AsInteger
      QWork.ParamByName("LAY").Value = CurrentQuery.FieldByName("TABELA").AsInteger'+#13#10+'  QWork.Active = True
      sDesc = QWork.FieldByName("PAI").AsString + "_" + QWork.FieldByName("TABELA").AsString
    Else
      QWork.Add("SELECT NOME FROM Z_TABELAS WHERE HANDLE = :HANDLE")
      QWork.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TABELA").AsInteger
      QWork.Active = True
      sDesc = QWork.FieldByName("NOME").AsString
    End If
    CurrentQuery.FieldByName("DESCRICAO").AsString = sDesc
    QWork.Active = False
    Set QWork = Nothing
  End If
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim obj As Object
  If CurrentQuery.FieldByName("ATUALIZAR").AsString = "N" Then
    Set obj = NewQuery
    obj.Add("SELECT * FROM Z_LEIAUTECAMPOSATUALIZAR WHERE LEIAUTE = :HANDLE")
    obj.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    obj.Active = True
    If Not obj.EOF Then
      bsShowMessage("Não é possivel desmarcar o campo ATUALIZAR, " + Chr(13) + "primeiramente deve ser excluído os campos a ignorar na atualização deste registro.", "I")
      CurrentQuery.FieldByName("ATUALIZAR").AsString = "S"
    End If
    obj.Active = False
    Set obj = Nothing
  End If
End Sub

