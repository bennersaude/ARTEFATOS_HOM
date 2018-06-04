'HASH: 96B58A4354A9B17C4797C8499AEE942F
'Macro: SAM_ESPECIALIDADEGRUPO_EXEC
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"
'#Uses "*RegistrarLogAlteracao"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T5674" Then
			EVENTOESTRUTURA.ReadOnly = True
			EVENTO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_AfterPost()
    RegistrarLogAlteracao "SAM_ESPECIALIDADEGRUPO_EXEC", CurrentQuery.FieldByName("HANDLE").AsInteger, "TABLE_AfterPost"
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
    RegistrarLogAlteracao "SAM_ESPECIALIDADEGRUPO", CurrentQuery.FieldByName("ESPECIALIDADEGRUPO").AsInteger, "SAM_ESPECIALIDADEGRUPO_EXEC.TABLE_BeforeDelete"
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If VisibleMode Then
		ESPECIALIDADEGRUPO.LocalWhere = "ORIGEMEVENTO = (SELECT ORIGEMEVENTO FROM SAM_TGE WHERE HANDLE = @EVENTO)"
	Else
	  ESPECIALIDADEGRUPO.WebLocalWhere = "ORIGEMEVENTO = (SELECT ORIGEMEVENTO FROM SAM_TGE WHERE HANDLE = @CAMPO(EVENTO))"
	End If

	If WebMode Then
		If WebVisionCode = "V_SAM_ESPECIALIDADEGRUPO_EXE_841" Then
			ESPECIALIDADE.ReadOnly = True
			ESPECIALIDADEGRUPO.ReadOnly = True
		End If
	End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If VisibleMode Then
		ESPECIALIDADEGRUPO.LocalWhere = "ORIGEMEVENTO = (SELECT ORIGEMEVENTO FROM SAM_TGE WHERE HANDLE = @EVENTO)"
	Else
	  ESPECIALIDADEGRUPO.WebLocalWhere = "ORIGEMEVENTO = (SELECT ORIGEMEVENTO FROM SAM_TGE WHERE HANDLE = @CAMPO(EVENTO))"
	End If

	If WebMode Then
		If WebVisionCode = "V_SAM_ESPECIALIDADEGRUPO_EXE_841" Then
			ESPECIALIDADE.ReadOnly = True
			ESPECIALIDADEGRUPO.ReadOnly = True
		End If
	End If

End Sub
