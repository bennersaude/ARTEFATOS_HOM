'HASH: FEA732985E30E379E448BB94C18F4E2B
'Macro: SFN_PESSOA_ISS
'#Uses "*bsShowMessage"


Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    COMPETENCIAFINAL.ReadOnly = False
  Else
    COMPETENCIAFINAL.ReadOnly = True
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    CanContinue = False
    bsShowMessage("Registro finalizado não pode ser alterado!", "E")
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SFN_PESSOA_ISS", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "PESSOA", "")

  If Linha <>"" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
  If CurrentQuery.FieldByName("TABRECOLHIMENTO").AsInteger = 1 Then
    CurrentQuery.FieldByName("MOTIVOISENCAO").Clear
    If(Not CurrentQuery.FieldByName("ALIQUOTA").IsNull)And(CurrentQuery.FieldByName("ALIQUOTA").AsInteger = 0)Then
    CanContinue = False
    bsShowMessage("Percentual=0 deve ser considerado como isento. Se for para considerar o percentual da cidade, deixe sem nenhum valor em percentual", "E")
    Exit Sub
  End If

Else
  CurrentQuery.FieldByName("ALIQUOTA").Clear

End If

End Sub

