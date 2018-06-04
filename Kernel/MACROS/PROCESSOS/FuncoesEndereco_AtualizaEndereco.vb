'HASH: C100298CEDD0C645FAA29303C4F869EE
Public Sub Main


    On Error GoTo erro:

	  Dim psOrigem As String
	  Dim piHandleOrigem As Long
	  Dim psXMLAtualizacao As String
	  Dim psXMLExclusao As String
	  Dim psMensagem As String
	  Dim piRetorno As Integer

  	  psOrigem = CStr( ServiceVar("psOrigem") )
	  piHandleOrigem = CLng( ServiceVar("piHandleOrigem") )
	  psXMLAtualizacao = CStr( ServiceVar("psXMLAtualizacao") )
	  psXMLExclusao = CStr( ServiceVar("psXMLExclusao") )

      Dim dllBSBen021    As Object
      Dim viHEndereco1 As Long
      Dim viHEndereco2 As Long
      Dim viHEndereco3 As Long
      Dim viHEndereco4 As Long

      Set dllBSBen021 = CreateBennerObject("BSBen021.AtualizacaoEndereco")

      If psXMLAtualizacao <> "" Then
      	psXMLAtualizacao = Replace( Replace( psXMLAtualizacao, "&lt", ">" ), "&gt", "<")
	  End If

	  If psXMLExclusao <> "" Then
	  	psXMLExclusao = Replace( Replace( psXMLExclusao, "&lt", ">" ), "&gt", "<")
	  End If

      If psOrigem = "B" Then
        piRetorno = dllBSBen021.Beneficiario(CurrentSystem, _
                                                                    piHandleOrigem, _
                                                                    False, _
                                                                    psXMLAtualizacao, _
                                                                    psXMLExclusao, _
                                                                    viHEndereco1, _
                                                                    viHEndereco2, _
                                                                    viHEndereco3, _
                                                                    viHEndereco4, _
                                                                    psMensagem)
      ElseIf psOrigem = "P" Then
        piRetorno = dllBSBen021.Pessoa(CurrentSystem, _
                                                              piHandleOrigem, _
                                                              False, _
                                                              psXMLAtualizacao, _
                                                              psXMLExclusao, _
                                                              viHEndereco1, _
                                                              viHEndereco2, _
                                                              psMensagem)
      ElseIf psOrigem = "C" Then
        piRetorno = dllBSBen021.Contrato(CurrentSystem, _
                                                               piHandleOrigem, _
                                                               False, _
                                                               psXMLAtualizacao, _
                                                               psXMLExclusao, _
                                                               viHEndereco1, _
                                                               psMensagem)
      End If

	  ServiceVar("psMensagem") = CStr( psMensagem )
	  ServiceVar("piRetorno") = CLng( piRetorno )

	  Set dllBSBen021 = Nothing

      Exit Sub

	erro:
		psMensagem = Err.Description
		piRetorno = 1
		ServiceVar("psMensagem") = CStr( psMensagem )
        ServiceVar("piRetorno") = CLng( piRetorno )

        Set dllBSBen021 = Nothing

End Sub
