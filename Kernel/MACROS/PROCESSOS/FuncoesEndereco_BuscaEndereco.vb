'HASH: D77FCDA8C46ECA1D4EE8EDA9408012F0

Public Sub Main
	On Error GoTo erro
		Dim psOrigem As String
		Dim piHandleOrigem As Long
		Dim psEnderecos As String
		Dim psMensagem As String
		Dim piRetorno As Integer
		Dim vsMensagem As String

		psOrigem = CStr( ServiceVar("psOrigem") )
		piHandleOrigem = CLng( ServiceVar("piHandleOrigem") )


		Dim especifico As Object
		Set especifico = CreateBennerObject("ESPECIFICO.UEspecifico")
	    vsMensagem = especifico.STF_BEN_VerificaPermissaoAlteracao(CurrentSystem, piHandleOrigem, psOrigem)
	    Set especifico = Nothing

		If vsMensagem <> "" Then
			psMensagem = vsMensagem
			piRetorno = 1
		Else
			 Dim dllBSBen021    As Object

		      Set dllBSBen021 = CreateBennerObject("BSBen021.BuscaEndereco")

		      If psOrigem = "B" Then
		        piRetorno = dllBSBen021.Beneficiario(CurrentSystem, _
		                                                                    piHandleOrigem, _
		                                                                    psEnderecos, _
		                                                                    psMensagem)
		      ElseIf psOrigem = "P" Then
		        piRetorno = dllBSBen021.Pessoa(CurrentSystem, _
		                                                              piHandleOrigem, _
		                                                              psEnderecos, _
		                                                              psMensagem)
		      ElseIf psOrigem = "C" Then
		        piRetorno = dllBSBen021.Contrato(CurrentSystem, _
		                                                               piHandleOrigem, _
		                                                               psEnderecos, _
		                                                               psMensagem)
		      End If

		End If

		ServiceVar("psEnderecos") = CStr( psEnderecos)
		ServiceVar("piRetorno") = CLng( piRetorno )
		ServiceVar("psMensagem") = CStr( psMensagem )

		Set dllBSBen021 = Nothing

		Exit Sub

	erro:
		ServiceVar("psMensagem") = Err.Description
		ServiceVar("piRetorno") = 1

		Set dllBSBen021 = Nothing

End Sub
