'HASH: 855BE2BB633F0A981EAC06F9B1FDB29E
'#Uses "*bsShowMessage"

Public Sub BOTAOCONSULTAR_OnClick()
  If CurrentQuery.State = 1 Then
    Dim Interface As Object
    Set Interface = CreateBennerObject("SFNTesouraria.Rotinas")
    Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger)
    Set Interface = Nothing
  Else
    bsShowMessage("Registro em Edição", "I")
  End If
End Sub

Public Sub BOTAOSALDOINICIAL_OnClick()
  If CurrentQuery.State = 1 Then

    Dim Interface As Object
	Dim vsMensagemErro As String
	Dim viRetorno As Integer
    Dim vvContainer As CSDContainer

	Set vvContainer = NewContainer

	Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")


	viRetorno = Interface.Exec(CurrentSystem, _
    						   1, _
                               "TV_FORM0037", _
                               "Lançamento de Saldo Inicial de Tesouraria", _
       	                       0, _
           	                   460, _
               	               490, _
                   	           False, _
                       	       vsMensagemErro, _
                           	   vvContainer)

	Select Case viRetorno
      Case -1
	  	bsShowMessage("Operação cancelada pelo usuário!", "I")
  	  Case  1
   	  	bsShowMessage(vsMensagemErro, "I")
	End Select

    Set Interface = Nothing

  Else
    bsShowMessage("Registro em Edição", "I")
  End If

End Sub

Public Sub BOTAOTRANSFERENCIA_OnClick()
  If CurrentQuery.State = 1 Then


    Dim Interface As Object
	Dim vsMensagemErro As String
	Dim viRetorno As Integer
    Dim vvContainer As CSDContainer

	Set vvContainer = NewContainer

	Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")


	viRetorno = Interface.Exec(CurrentSystem, _
    						   1, _
                               "TV_FORM0039", _
                               "Transferencia entre Tesourarias", _
       	                       0, _
           	                   460, _
               	               490, _
                   	           False, _
                       	       vsMensagemErro, _
                           	   vvContainer)

	Select Case viRetorno
      Case -1
	  	bsShowMessage("Operação cancelada pelo usuário!", "I")
  	  Case  1
   	  	bsShowMessage(vsMensagemErro, "I")
	End Select

    Set Interface = Nothing
  Else
    bsShowMessage("Registro em Edição", "I")
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim ParamFin As Object
    Set ParamFin = NewQuery

    ParamFin.Active = False
    ParamFin.Add("SELECT UTILIZATESOURARIA, CONTABILIZA FROM SFN_PARAMETROSFIN")
    ParamFin.Active = True

    If ParamFin.FieldByName("CONTABILIZA").AsString = "S" Then
      If CurrentQuery.FieldByName("HISTORICOTRANSFERENCIA").IsNull Then
        bsShowMessage("Histórico padrão obrigatório", "E")
        CanContinue = False
      End If
      If CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        bsShowMessage("Classe contábil obrigatório", "E")
        CanContinue = False
      End If
    End If
    Set ParamFin = Nothing
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCONSULTAR"
			BOTAOCONSULTAR_OnClick
	End Select
End Sub
