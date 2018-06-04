'HASH: 4B6898931D0A76EE2F56606767187015
'Macro: SAM_GRUPOCONTRATO


Public Sub BOTAOIMPORTAR_OnClick()
  Dim OLEImp As Object
  Set OLEImp = CreateBennerObject("SamImpSal.Salario")
  OLEImp.Exec(CurrentSystem, 0)
  Set OLEImp = Nothing
End Sub

Public Sub BOTAOOCORRENCIAS_OnClick()
  Dim OLEImp As Object
  Set OLEImp = CreateBennerObject("SamImpSal.Salario")
  OLEImp.Ocorrencia(CurrentSystem)
  Set OLEImp = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  'Daniela -SMS 12220 -Convênio no registro da ANS
  If Not CurrentQuery.FieldByName("CONVENIO").IsNull Then
    CONVENIO.ReadOnly = True
  Else
    CONVENIO.ReadOnly = False
  End If
  If WebMode Then
  	If WebVisionCode = "V_SAM_GRUPOCONTRATO" Then
  		CONVENIO.ReadOnly = True
  	End If
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOIMPORTAR"
			BOTAOIMPORTAR_OnClick
 		Case "BOTAOOCORRENCIAS"
 			BOTAOOCORRENCIAS_OnClick
	End Select
End Sub
