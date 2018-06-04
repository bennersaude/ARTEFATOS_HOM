'HASH: 8D5919647589D7E65F82107E9A43773B

Public Sub BOTAOCANCELAR_OnClick()
  Dim sp_CancelaSibXML As BStoredProc

  Set sp_CancelaSibXML = NewStoredProc

  sp_CancelaSibXML.Name="BSANS_SIBENVIO_CANCXML"
  sp_CancelaSibXML.AddParam("p_RotXml",ptInput, ftInteger,4)
  sp_CancelaSibXML.ParamByName("p_RotXml").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sp_CancelaSibXML.ExecProc

  Set sp_CancelaSibXML = Nothing

  If Not WebMode Then
  	RefreshNodesWithTable("ANS_SIBPADRAO_XML")
  End If

End Sub
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
 		Case "BOTAOCANCELAR"
 			BOTAOCANCELAR_OnClick
	End Select
End Sub
