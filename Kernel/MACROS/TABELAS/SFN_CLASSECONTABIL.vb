'HASH: 880FA888624949BF60805498884AFF57
 
'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)  'SMS 131975 - GUSTAVO GALINA - 10/06/2010
Dim qClasseContabil As BPesquisa
Set qClasseContabil = NewQuery

qClasseContabil.Clear
qClasseContabil.Add("SELECT HANDLE                   ")
qClasseContabil.Add("  FROM SFN_CLASSECONTABIL       ")
qClasseContabil.Add(" WHERE CODIGOREDUZIDO = :CODIGO ")
qClasseContabil.Add("   AND HANDLE <> :HANDLEATUAL ")

qClasseContabil.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGOREDUZIDO").AsInteger
qClasseContabil.ParamByName("HANDLEATUAL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
qClasseContabil.Active = True

If qClasseContabil.FieldByName("HANDLE").AsInteger > 0 Then
	BsShowMessage("Esse código reduzido já existe, por favor digite outro!", "E")
    If VisibleMode Then
	  CODIGOREDUZIDO.SetFocus
	End If
	CanContinue = False
End If

Set qClasseContabil= Nothing
End Sub
