'HASH: 3DE44FFBA3F43C824131A296E324ECE7
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If((CurrentQuery.FieldByName("CHECKCONTRATO").AsString ="N")Or(CurrentQuery.FieldByName("CHECKCONTRATO").AsString =""))And((CurrentQuery.FieldByName("CHECKFAMILIA").AsString = "N")Or(CurrentQuery.FieldByName("CHECKFAMILIA").AsString = ""))Then
   		bsShowMessage("É necessário selecionar pelo menos um dos itens abaixo para a impressão:" +Chr(13)+"- Módulo contrato" +Chr(13)+"- Família", "E")
   		CanContinue = False
	End If
End Sub
