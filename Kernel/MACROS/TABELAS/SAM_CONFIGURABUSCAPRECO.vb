'HASH: 45E7B9C260E22BF2061815D0F8E871D3
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vNivel1 As Integer
  Dim vNivel2 As Integer
  Dim vNivel3 As Integer
  Dim vNivel4 As Integer
  Dim vNivel5 As Integer

  vNivel1 = CurrentQuery.FieldByName("NIVEL1").AsInteger
  vNivel2 = CurrentQuery.FieldByName("NIVEL2").AsInteger
  vNivel3 = CurrentQuery.FieldByName("NIVEL3").AsInteger
  vNivel4 = CurrentQuery.FieldByName("NIVEL4").AsInteger
  vNivel5 = CurrentQuery.FieldByName("NIVEL5").AsInteger

  If (vNivel1 = vNivel2 Or vNivel1 = vNivel3 Or vNivel1 = vNivel4 Or vNivel1 = vNivel5) And Not (CurrentQuery.FieldByName("NIVEL1").IsNull) Then
    CanContinue = False
    bsShowMessage("Não pode existir dois niveis iguais !", "E")
    Exit Sub
  End If
  If (vNivel2 = vNivel1 Or vNivel2 = vNivel3 Or vNivel2 = vNivel4 Or vNivel2 = vNivel5) And Not (CurrentQuery.FieldByName("NIVEL2").IsNull) Then
    CanContinue = False
    bsShowMessage("Não pode existir dois niveis iguais !", "E")
    Exit Sub
  End If
  If (vNivel3 = vNivel1 Or vNivel3 = vNivel2 Or vNivel3 = vNivel4 Or vNivel3 = vNivel5) And Not (CurrentQuery.FieldByName("NIVEL3").IsNull) Then
    CanContinue = False
    bsShowMessage("Não pode existir dois niveis iguais !", "E")
    Exit Sub
  End If
  If (vNivel4 = vNivel1 Or vNivel4 = vNivel2 Or vNivel4 = vNivel3 Or vNivel4 = vNivel5) And Not (CurrentQuery.FieldByName("NIVEL4").IsNull) Then
    CanContinue = False
    bsShowMessage("Não pode existir dois niveis iguais !", "E")
    Exit Sub
  End If
  If (vNivel5 = vNivel1 Or vNivel5 = vNivel2 Or vNivel5 = vNivel3 Or vNivel5 = vNivel4) And Not (CurrentQuery.FieldByName("NIVEL5").IsNull) Then
    CanContinue = False
    bsShowMessage("Não pode existir dois niveis iguais !", "E")
    Exit Sub
  End If
End Sub

