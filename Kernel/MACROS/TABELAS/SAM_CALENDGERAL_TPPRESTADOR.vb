'HASH: 91BDB0CC69447E4C36786B4F1E79E565
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("TIPOPRESTADOR").IsNull Then
    bsShowMessage("Campo tipo prestador é obrigatório.", "E")
    CanContinue = False
    Exit Sub
  End If


  '  Dim Q As Object
  '  Set Q = NewQuery
  '  Q.Add("SELECT TIPOPRESTADOR FROM SAM_CALENDGERAL_TPPRESTADOR WHERE TIPOPRESTADOR IS NULL AND HANDLE <> :HANDLE")
  '  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("handle").AsInteger
  '  Q.Active = True
  '  If Not Q.EOF  Then
  '    MsgBox "Já existe pagamento com condição Nulo - Somente é permitido um único tipo de prestador nulo."
  '    CanContinue = False
  '  End If

  'End If

  '  If Not CurrentQuery.FieldByName("TIPOPRESTADOR").IsNull Then

  Dim Q2 As Object
  Set Q2 = NewQuery
  Q2.Add("SELECT TIPOPRESTADOR FROM SAM_CALENDGERAL_TPPRESTADOR WHERE TIPOPRESTADOR = :TIPOPRESTADOR AND HANDLE <> :HANDLE")
  Q2.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("handle").AsInteger
  Q2.ParamByName("TIPOPRESTADOR").Value = CurrentQuery.FieldByName("TIPOPRESTADOR").AsInteger
  Q2.Active = True
  If Not Q2.EOF Then
    bsShowMessage("Este tipo de prestador já está no calendário geral.", "E")
    CanContinue = False
  End If

  '  End If

End Sub

