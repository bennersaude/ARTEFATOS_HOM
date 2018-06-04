'HASH: 818D92F1FA633A905AC4C349C10BA6FF
 
'#Uses "*bsShowMessage

Public Sub TABLE_AfterInsert()
    CurrentQuery.FieldByName("REGRAFINANCEIRA").AsInteger = 1
    CurrentQuery.FieldByName("VALOR").AsInteger = 1
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim dllFaturaAvulsa As Object
	Dim vMensagemRetorno As String


    On Error GoTo Erro
    vMensagemRetorno = vMensagemRetorno + " 2 "

'    Set dllFaturaAvulsa = CreateBennerObject("SFNFATURA.Rotinas")
'    vMensagemRetorno = vMensagemRetorno + dllFaturaAvulsa.FaturaAvulsaWeb(CurrentSystem, vContaFin, vTipoFatura, xmlFaturaAvulsa)


   Dim vdata As Date
    Set dllFaturaAvulsa = CreateBennerObject("SFNBAIXA.Documento")
    vMensagemRetorno = vMensagemRetorno + dllFaturaAvulsa.BxDocWeb(CurrentSystem, 217037, -1, vdata, vdata, "-1", -1, -1, "motivo", -1, -2, -3, -4, -5)


    Set dllFaturaAvulsa = Nothing

     GoTo Fim

    Erro :
    vMensagemRetorno = vMensagemRetorno + Error


    Fim :
	bsShowMessage(vMensagemRetorno, "I")

End Sub
