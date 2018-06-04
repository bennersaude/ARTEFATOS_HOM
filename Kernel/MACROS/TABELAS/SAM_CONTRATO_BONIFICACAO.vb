'HASH: C1FE44CDF15ABB624C4E42A11EF6C1DB
'Macro: SAM_CONTRATO_BONIFICACAO
'#Uses "*bsShowMessage"


Public Sub TABLE_AfterScroll()
  If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    COMPETENCIAFINAL.ReadOnly = True
  Else
    COMPETENCIAFINAL.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If(Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And _
     (CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime <CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)Then
  bsShowMessage("A Competência final, se informada, deve ser maior ou igual a inicial", "E")
  CanContinue = False
Else
  CanContinue = True
End If
Dim Interface As Object
Dim Linha As String

Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_BONIFICACAO", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "CONTRATO", "")

If Linha = "" Then
  CanContinue = True
Else
  CanContinue = False
  bsShowMessage(Linha, "E")
End If

End Sub

