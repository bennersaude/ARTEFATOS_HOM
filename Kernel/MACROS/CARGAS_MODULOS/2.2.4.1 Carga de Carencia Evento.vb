'HASH: ABC69C54B551314ECE58CBDD51CCB3F0
Public Sub BOTAOGERAEVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica =CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem,"SAM_CARENCIA_EVENTO","CARENCIA",RecordHandleOfTable("SAM_CARENCIA"),"Duplicando eventos para carência")
  Set Duplica =Nothing
  RefreshNodesWithTable "SAM_CARENCIA_EVENTO"
End Sub
