'HASH: EDBCCCAE357C3751C1487D92DE2E30C2
Option Explicit

'#uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If Not ((CurrentQuery.FieldByName("CODIGOGLOSATISS").AsInteger >= 9901) And _
	        (CurrentQuery.FieldByName("CODIGOGLOSATISS").AsInteger <= 9999)) Then

		bsShowMessage("Não é possível excluir registros fora do intervalo permitido (9901 a 9999).", "E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If Not ((CurrentQuery.FieldByName("CODIGOGLOSATISS").AsInteger >= 9901) And _
	        (CurrentQuery.FieldByName("CODIGOGLOSATISS").AsInteger <= 9999)) Then
		bsShowMessage("Não é possível editar registros fora do intervalo permitido (9901 a 9999).", "E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If Not ((CurrentQuery.FieldByName("CODIGOGLOSATISS").AsInteger >= 9901) And _
	        (CurrentQuery.FieldByName("CODIGOGLOSATISS").AsInteger <= 9999)) Then
		bsShowMessage("Não é possível inserir registros fora do intervalo permitido (9901 a 9999).", "E")
		CanContinue = False
	End If
End Sub
