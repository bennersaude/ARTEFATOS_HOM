'HASH: 33B6EC7D74C7EE75359F9127FF817A9C
 
Option Explicit

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID="BOTAOREIMPRIMIR" Then
		Dim retorno As Integer
		Dim mensagem As String
		Dim dll As Object
		Set dll=CreateBennerObject("samauto.autorizador")
		retorno = dll.reimprimir(CurrentSystem, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, CurrentQuery.FieldByName("NUMEROGUIA").AsInteger, mensagem)
		Set dll=Nothing
		If retorno>0 Then
			InfoDescription = mensagem
		Else
			InfoDescription = "Reimpressão concluída com sucesso"
		End If
	End If
End Sub
