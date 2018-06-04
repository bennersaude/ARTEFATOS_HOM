'HASH: D77E9A5B0A3B0A5E21A4672C61DF3C48
'Macro: SAM_AGENCIACOMISSAO
 

Public Sub TABLE_AfterScroll()
        ROTULORESPONSAVEL.Text= ""

	    Dim SQL As Object
	    
	    Set SQL = NewQuery


		    SQL.Add("SELECT C.TABRESPONSAVEL, B.NOME BENEFNOME, P.NOME PESNOME")
		    SQL.Add("FROM SFN_CONTAFIN C")
		    SQL.Add("     LEFT JOIN SAM_BENEFICIARIO B ON")
	    	SQL.Add("     (B.HANDLE = C.BENEFICIARIO)")
	    	SQL.Add("     LEFT JOIN SFN_PESSOA P ON")
	    	SQL.Add("     (P.HANDLE = C.PESSOA)")
	    	SQL.Add("WHERE C.HANDLE = :HCONTAFIN")
		SQL.ParamByName("HCONTAFIN").Value=CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
	    SQL.Active=True
	    
	    If SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
			ROTULORESPONSAVEL.Text="Responsável: " + SQL.FieldByName("BENEFNOME").AsString
		ElseIf SQL.FieldByName("TABRESPONSAVEL").AsInteger = 3 Then
			ROTULORESPONSAVEL.Text="Responsável: " + SQL.FieldByName("PESNOME").AsString
		End If
	    
	    SQL.Active=False
	    Set SQL = Nothing

End Sub
