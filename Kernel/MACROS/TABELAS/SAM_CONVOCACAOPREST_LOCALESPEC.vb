'HASH: BFE541F447BCD3152EAAC9403A55B832
Option Explicit

Dim VAGAILIMITADA As Boolean


Public Sub TABLE_AfterScroll()
	VAGAILIMITADA = CurrentQuery.FieldByName("VAGAILIMITADA").AsBoolean
	If VAGAILIMITADA Then
		VAGADISPONIVEL.ReadOnly = True
	Else
		VAGADISPONIVEL.ReadOnly = False
	End If
End Sub

Public Sub VAGAILIMITADA_OnChange()
	VAGAILIMITADA = Not VAGAILIMITADA
	If VAGAILIMITADA Then
		CurrentQuery.FieldByName("VAGADISPONIVEL").AsInteger = 999999
		VAGADISPONIVEL.ReadOnly = True
	Else
	    CurrentQuery.FieldByName("VAGADISPONIVEL").AsInteger = 1
		VAGADISPONIVEL.ReadOnly = False
	End If
End Sub
