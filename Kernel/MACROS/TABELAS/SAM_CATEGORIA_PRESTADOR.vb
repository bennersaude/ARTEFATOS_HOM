'HASH: 0A324BD3C0298948AA0E516F0A4C455E

'MACRO: SAM_CATEGORIA_PRESTADOR
'#Uses "*bsShowMessage"

Public Function VerificaCPFDuplicado(handleCategoriaPrestador As Integer, tipoValidacao As String) As Boolean

	Dim interface As CSEntityCall

	If tipoValidacao = "titular" Then
	    Set interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfBeneficiarioTitularIgualCpfPrestador")

	ElseIf tipoValidacao = "dependente" Then
		Set interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfBeneficiarioDependenteIgualCpfPrestador")

	ElseIf tipoValidacao = "usuario" Then
		Set interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfUsuarioIgualCpfPrestador")

	End If

	interface.AddParameter(pdtInteger, handleCategoriaPrestador)

	VerificaCPFDuplicado = CBool(interface.Execute())

	Set interface = Nothing

End Function


Public Sub TABLE_AfterDelete()
  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SAM_CATEGORIA_PRESTADOR")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "Z")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_AfterPost()
  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SAM_CATEGORIA_PRESTADOR")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim mensagem As String

  On Error GoTo ERRO:

  If CurrentQuery.FieldByName("BLOQUEARINCLUSAOBENEF").AsString = "S" Then

    If VerificaCPFDuplicado(CurrentQuery.FieldByName("HANDLE").AsInteger, "titular") Then

		mensagem = "Existe prestador da categoria com o mesmo CPF de beneficiário ativo." + vbCrLf

    Else

	  If VerificaCPFDuplicado(CurrentQuery.FieldByName("HANDLE").AsInteger, "dependente") Then

        mensagem = "Existe prestador da categoria com o mesmo CPF de beneficiário ativo." + vbCrLf

      End If

	End If

    If VerificaCPFDuplicado(CurrentQuery.FieldByName("HANDLE").AsInteger, "usuario") Then

    	mensagem = mensagem + "Existe prestador da categoria com o mesmo CPF de usuário do sistema." + vbCrLf

    End If

    If mensagem <> "" Then

	    If bsShowMessage(mensagem + "Deseja continuar?", "Q") = vbNo Then
	      If VisibleMode Then
	        CanContinue = False
	      End If
	      Exit Sub
	    End If

    End If

  End If

  ERRO:
    If Error <> "" Then
	  bsShowMessage(Error, "E")
      CanContinue = False
	End If

End Sub
