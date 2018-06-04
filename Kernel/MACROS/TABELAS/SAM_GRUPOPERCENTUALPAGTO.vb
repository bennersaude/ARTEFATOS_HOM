'HASH: 0B1BA68EDABE4A6E8CB6FF6F8DDF089D

Public Sub BOTAGERAREVENTO_OnClick()
  Dim Duplica As Object
  Set Duplica = CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem, "SAM_GRUPOPERCENTUALPAGTO_EVENT", "GRUPOPERCENTUALPAGTO", RecordHandleOfTable("SAM_GRUPOPERCENTUALPAGTO"), "Gerando eventos para Grupos de Percentuais de Pagamento")
  Set Duplica = Nothing
  RefreshNodesWithTable"SAM_GRUPOPERCENTUALPAGTO_EVENT"

End Sub



Public Sub TABLE_AfterScroll()
	If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    	BOTAGERAREVENTO.Enabled = True
	  Else
	    BOTAGERAREVENTO.Enabled = False
	End If
End Sub
