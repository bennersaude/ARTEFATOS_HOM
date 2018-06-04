'HASH: A002FF3E520D2C06D32462BCCC8B7194
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    COMPETENCIAFINAL.ReadOnly = False
  Else
    COMPETENCIAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    If CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime > _
                                CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime Then
      CanContinue = False
      bsShowMessage("A competência final não pode ser inferior a competência inicial", "E")
      Exit Sub
    End If
  End If

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_MODDESCFAIXA", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "CONTRATOMOD", Condicao)

  If Linha <>"" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing

End Sub

