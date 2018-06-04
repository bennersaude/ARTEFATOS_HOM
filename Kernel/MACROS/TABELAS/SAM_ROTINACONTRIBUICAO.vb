'HASH: BD6E8569B1EAAE22DDB0766B2F9EE895
 
'Macro SAM_ROTINACONTRIBUICAO (SMS 60044)



Public Sub BOTAOCANCELAR_OnClick()
  Dim Interface As Object

  Set Interface = CreateBennerObject("BSDIV001.ROTINAS")
  Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "C")

  Set Interface = Nothing

  WriteAudit("C", HandleOfTable("SAM_ROTINACONTRIBUICAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Contribuição - Cancelamento de Processamento")

End Sub

Public Sub BOTAOCANCELARIMPORTACAO_OnClick()

  Dim Interface As Object

  Set Interface = CreateBennerObject("BSDIV001.ROTINAS")
  Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "D")

  Set Interface = Nothing

  WriteAudit("D", HandleOfTable("SAM_ROTINACONTRIBUICAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Contribuição - Cancelamento de Importação")

End Sub

Public Sub BOTAOIMPORTAR_OnClick()

  Dim Interface As Object

  Set Interface = CreateBennerObject("BSDIV001.ROTINAS")
  Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "I")

  Set Interface = Nothing

  WriteAudit("I", HandleOfTable("SAM_ROTINACONTRIBUICAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Contribuição - Importação")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Interface As Object

  Set Interface = CreateBennerObject("BSDIV001.ROTINAS")
  Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "P")

  Set Interface = Nothing

  WriteAudit("P", HandleOfTable("SAM_ROTINACONTRIBUICAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Contribuição - Processamento")

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOCANCELARIMPORTACAO"
			BOTAOCANCELARIMPORTACAO_OnClick
		Case "BOTAOIMPORTAR"
			BOTAOIMPORTAR_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
