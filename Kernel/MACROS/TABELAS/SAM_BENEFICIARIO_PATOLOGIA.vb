'HASH: FF2BBAEF7688DFF357334D1A2BA6FFB3
 
'#Uses "*bsShowMessage"
Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "SAM_BENEFICIARIO_PATOLOGIA", "Duplicando Patologias para Beneficiário", "SAM_PATOLOGIA", "PATOLOGIA", "BENEFICIARIO", CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "N", "")
  Set Obj = Nothing

End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Excluir(CurrentSystem, "SAM_BENEFICIARIO_PATOLOGIA", "Excluindo Patologias para Beneficiário", "SAM_PATOLOGIA", "PATOLOGIA", "BENEFICIARIO", CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "N", "")
  Set Obj = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim QTemp As Object
    Set QTemp = NewQuery

    QTemp.Active = False
    QTemp.Clear
    QTemp.Add("SELECT COUNT(HANDLE) QT ")
    QTemp.Add("  FROM SAM_BENEFICIARIO_PATOLOGIA ")
    QTemp.Add(" WHERE BENEFICIARIO = :BEN")
    QTemp.Add("   AND PATOLOGIA = :PAT")
    QTemp.ParamByName("BEN").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    QTemp.ParamByName("PAT").AsInteger = CurrentQuery.FieldByName("PATOLOGIA").AsInteger
    QTemp.Active = True
    If QTemp.FieldByName("QT").AsInteger > 0 Then
      bsShowMessage("Registro já cadatrado!", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    bsShowmessage("Se for restrição de plano antigo não regulamentado, altere o registro preenchendo data final!", "I")
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOGERAR"
			BOTAOGERAR_OnClick
		Case "BOTAOEXCLUIR"
			BOTAOEXCLUIR_OnClick
	End Select
End Sub
