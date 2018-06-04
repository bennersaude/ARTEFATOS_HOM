'HASH: 67F4865D869335144CB71CEC3282AF19

'#Uses "*bsShowMessage"

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  '#Uses "*ProcuraPrestador"

  'Dim vHandle As Long
  'ShowPopup = False
  'vHandle = ProcuraPrestador("","T", PRESTADOR.Text)  ' pelo CPF e recebedor
  'If vHandle<>0 Then
  '   CurrentQuery.Edit
  '   CurrentQuery.FieldByName("PRESTADOR").Value=vHandle
  'End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
    bsShowMessage("Campo Prestador é obrigatório.", "E")
    CanContinue = False
    Exit Sub
  End If

  '  Dim Q As Object
  '  Set Q = NewQuery
  '  Q.Add("SELECT PRESTADOR FROM SAM_CALENDGERAL_PRESTADOR WHERE PRESTADOR IS NULL AND HANDLE <> :HANDLE")
  '  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  '  Q.Active = True
  '  If Not Q.EOF  Then
  '    MsgBox "Já existe pagamento com condição Nulo - Somente é permitido um único prestador nulo."
  '    CanContinue = False
  '  End If

  'End If


  'If Not CurrentQuery.FieldByName("PRESTADOR").IsNull Then

  Dim Q2 As Object
  Set Q2 = NewQuery
  Q2.Add("SELECT PRESTADOR FROM SAM_CALENDGERAL_PRESTADOR WHERE PRESTADOR = :PRESTADOR AND HANDLE <> :HANDLE")
  Q2.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q2.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  Q2.Active = True
  If Not Q2.EOF Then
    bsShowMessage("Este prestador já está no calendário geral.", "E")
    CanContinue = False
  End If

  'End If



End Sub

