'HASH: C72486CF0154462C269B115B82EF058A
 

Public Sub TABLE_AfterScroll()

	If VisibleMode Then
	  ROTPRECOGENERICO.Visible = False
	ElseIf WebMode Then

	  If CurrentQuery.FieldByName("PRECOGENERICO").AsInteger > 0 Then
		Dim SQL As Object
		Set SQL = NewQuery
		SQL.Add("SELECT QTDUSHONORARIO FROM SAM_PRECOGENERICO_DOTAC WHERE HANDLE = :HANDLE")
		SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PRECOGENERICO").AsInteger
		SQL.Active = True

        ROTPRECOGENERICO.Text = "Preço genérico: " + Format(SQL.FieldByName("QTDUSHONORARIO").AsFloat,"###,###,##0.0000")

        Set SQL = Nothing
	  End If
	End If

End Sub
