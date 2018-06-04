'HASH: 481FA282EDA3775006716DFCA531F1A9
Public Sub TABLE_AfterScroll()
	TIPOCREDENCIAMENTO.WebLocalWhere = "HANDLE IN (SELECT TIPOCREDENCIAMENTO FROM SAM_PRESTADOR_PROC_CREDEN WHERE PRESTADORPROCESSO = " + CStr(RecordHandleOfTable("SAM_PRESTADOR_PROC")) + ")"

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim OLESamCarta As Object
  Set OLESamCarta = CreateBennerObject("SamCarta.Impressao")
  OLESamCarta.imprimirCarta(CurrentSystem, RecordHandleOfTable("SAM_PRESTADOR_PROC"), CurrentQuery.FieldByName("TIPOCREDENCIAMENTO").AsInteger)
  Set OLESamCarta = Nothing
End Sub
