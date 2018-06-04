'HASH: DCA5A2B2DABBA933BD2BF0D69883019D
 Option Explicit
Public Sub FATURA_OnChange()
	Dim qFatura As BPesquisa
  	Set qFatura = NewQuery

	qFatura.Clear
  	qFatura.Add("SELECT SITUACAO   ")
  	qFatura.Add("  FROM SFN_FATURA ")
  	qFatura.Add(" WHERE HANDLE=:PFATURA")
  	qFatura.ParamByName("PFATURA").AsInteger = CurrentQuery.FieldByName("FATURA").AsInteger
  	qFatura.Active = True

  	If qFatura.FieldByName("SITUACAO").AsString = "A" Then
  		CurrentQuery.FieldByName("SITUACAO").AsString = "Aberta"
  	ElseIf qFatura.FieldByName("SITUACAO").AsString = "B" Then
		CurrentQuery.FieldByName("SITUACAO").AsString = "Baixada"
	ElseIf qFatura.FieldByName("SITUACAO").AsString = "C" Then
		CurrentQuery.FieldByName("SITUACAO").AsString = "Cancelada"
	ElseIf qFatura.FieldByName("SITUACAO").AsString = "S" Then
		CurrentQuery.FieldByName("SITUACAO").AsString = "Suspensa"
  	End If
  	Set qFatura = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  If VisibleMode Then
		FATURA.LocalWhere = "EXISTS( SELECT 1 FROM SFN_FATURA_LANC LANC " + _
		                    "JOIN SFN_TIPOLANCFIN C ON (C.HANDLE = LANC.TIPOLANCFIN) " + _
                            "JOIN SIS_TIPOLANCFIN D ON (D.HANDLE = C.TIPOLANCFIN) " + _
                            "JOIN SFN_CONTAFIN CF ON (CF.HANDLE = A.CONTAFINANCEIRA) " + _
                            "JOIN SAM_BENEFICIARIO BENEF ON (BENEF.HANDLE = CF.BENEFICIARIO) " + _
                            "JOIN SAM_FAMILIA FAM ON (FAM.HANDLE = BENEF.FAMILIA) " + _
                            "JOIN SAM_BENEFICIARIO BENEF2 ON (BENEF2.FAMILIA = FAM.HANDLE) " + _
                            "WHERE LANC.FATURA = A.HANDLE AND D.CODIGO = 61 " + _
                            "AND BENEF2.HANDLE = " + CStr(RecordHandleOfTable("SAM_BENEFICIARIO")) + ")"
	ElseIf WebMode Then
		FATURA.WebLocalWhere = "EXISTS( SELECT 1 FROM SFN_FATURA_LANC LANC " + _
		                       "JOIN SFN_TIPOLANCFIN C ON (C.HANDLE = LANC.TIPOLANCFIN) " + _
                               "JOIN SIS_TIPOLANCFIN D ON (D.HANDLE = C.TIPOLANCFIN) " + _
                               "JOIN SFN_CONTAFIN CF ON (CF.HANDLE = A.CONTAFINANCEIRA) " + _
                               "JOIN SAM_BENEFICIARIO BENEF ON (BENEF.HANDLE = CF.BENEFICIARIO) " + _
                               "JOIN SAM_FAMILIA FAM ON (FAM.HANDLE = BENEF.FAMILIA) " + _
                               "JOIN SAM_BENEFICIARIO BENEF2 ON (BENEF2.FAMILIA = FAM.HANDLE) " + _
                               "WHERE LANC.FATURA = A.HANDLE AND D.CODIGO = 61 " + _
                               "AND BENEF2.HANDLE = " + CStr(RecordHandleOfTable("SAM_BENEFICIARIO")) + ")"
	End If
End Sub
