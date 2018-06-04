'HASH: 4389CDFE0681953EB1C599EDADE10F8F
'Macro: SAM_MIGRACAO_BENEF_PFEVENTO
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = " AND TABELAPFEVENTO = " + CurrentQuery.FieldByName("TABELAPFEVENTO").AsString
  Condicao = Condicao + " AND TABTIPOPF = " + CurrentQuery.FieldByName("TABTIPOPF").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_PFEVENTO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

End Sub

Public Sub TABTIPOPF_OnChange()
  If CurrentQuery.State = 3 Then
    If CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 1 Then
      CurrentQuery.FieldByName("PERIODO").Value = Null
    Else
      CurrentQuery.FieldByName("CODIGOPF").Value = Null
    End If
  End If
End Sub

