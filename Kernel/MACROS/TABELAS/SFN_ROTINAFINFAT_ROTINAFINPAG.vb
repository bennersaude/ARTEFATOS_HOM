'HASH: 334545E317D7ACD6831C8660AABC36AA
 
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("TIPOFATURAMENTO").Clear
	CurrentQuery.FieldByName("COMPETFIN").Clear
	CurrentQuery.FieldByName("ROTINAFIN").Clear
End Sub


Public Sub VerificaSeProcessada(CanContinue As Boolean)

  Dim HandleRotinaFinFat As Long
  HandleRotinaFinFat = RecordHandleOfTable("SFN_ROTINAFINFAT")
  If (HandleRotinaFinFat > 0) Then
  	Dim SQLRotFin As Object
  	Set SQLRotFin = NewQuery
  	SQLRotFin.Add("SELECT A.SITUACAOFATURAMENTO FROM SFN_ROTINAFINFAT A")
  	SQLRotFin.Add("WHERE A.HANDLE = :ROTINAFINFAT")
  	SQLRotFin.ParamByName("ROTINAFINFAT").AsInteger = HandleRotinaFinFat
  	SQLRotFin.Active = True
  	If SQLRotFin.FieldByName("SITUACAOFATURAMENTO").AsString <> "1" Then
	    CanContinue = False
	    bsShowMessage("A Rotina já foi processada!", "E")
  	End If
  	SQLRotFin.Active = False
  	Set SQLRotFin = Nothing
  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub
