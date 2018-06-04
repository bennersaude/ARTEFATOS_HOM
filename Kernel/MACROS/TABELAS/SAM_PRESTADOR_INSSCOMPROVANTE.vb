'HASH: FD69F2CD706EA0CB79123513974A69CC
'MACRO SAM_PRESTADOR_INSSCOMPROVANTE
'#Uses "*CheckPrestador"
'#Uses "*ShowMsg"
'#Uses "*bsShowMessage"

Option Explicit

Dim Component As CSBusinessComponent

Public Sub CODIGOCATEGORIATRABALHADOR_OnPopup(ShowPopup As Boolean)
  CreateComponent

  CODIGOCATEGORIATRABALHADOR.LocalWhere = Component.Execute("FiltrarCategoriasTrabalhador")

  DestroyComponent
End Sub

Public Sub TABLE_AfterDelete()
  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SAM_PRESTADOR_INSSCOMPROVANTE")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "Z")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_AfterPost()
  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SAM_PRESTADOR_INSSCOMPROVANTE")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_AfterScroll()
	Dim CNPJCPF As String

	CNPJCPF = CurrentQuery.FieldByName("CNPJCPF").AsString

	If Len(CNPJCPF) = 11 Then
		CurrentQuery.FieldByName("CNPJCPF").Mask = "999\.999\.999\-99;0;_"
	ElseIf Len(CNPJCPF) = 14 Then
		CurrentQuery.FieldByName("CNPJCPF").Mask = "99\.999\.999\/9999\-99;0;_"
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If (ExisteVinculoESocial(CurrentQuery.FieldByName("HANDLE").AsInteger)) Then
    bsShowMessage("Comprovante não pode ser removido, pois está vinculado ao eSocial.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
    If (ExisteVinculoESocial(CurrentQuery.FieldByName("HANDLE").AsInteger)) Then
      bsShowMessage("Comprovante não pode ser alterado, pois está vinculado ao eSocial.", "E")
      CanContinue = False
      Exit Sub
    End If

	Dim CNPJCPF As String

	CNPJCPF = CurrentQuery.FieldByName("CNPJCPF").AsString

	If Len(CNPJCPF) = 11 Then
		CurrentQuery.FieldByName("CNPJCPF").Mask = "999\.999\.999\-99;0;_"
	ElseIf Len(CNPJCPF) = 14 Then
		CurrentQuery.FieldByName("CNPJCPF").Mask = "99\.999\.999\/9999\-99;0;_"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim CNPJCPF As String

	If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
		If CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime < CurrentQuery.FieldByName("COMPETENCIA").AsDateTime Then
			bsShowMessage("A competência final não pode ser menor que a inicial!", "E")
			CanContinue = False
			Exit Sub
		End If
	End If

	CNPJCPF = CurrentQuery.FieldByName("CNPJCPF").AsString

	If Len(CNPJCPF) = 11 Then
		If Not IsValidCPF(CNPJCPF) Then
			bsShowMessage("CPF Inválido", "E")
			CanContinue = False
			Exit Sub
		End If
	ElseIf Len(CNPJCPF) = 14 Then
		If Not IsValidCGC(CNPJCPF) Then
			bsShowMessage("CNPJ Inválido", "E")
			CanContinue = False
			Exit Sub
		End If
	Else
		bsShowMessage("CPF / CNPJ Inválido", "E")
		CanContinue = False
		Exit Sub
	End If

	If ((Not CurrentQuery.FieldByName("CODIGOCATEGORIATRABALHADOR").IsNull) Or (Not CurrentQuery.FieldByName("INDMV").IsNull)) Then
      CreateComponent

      Component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
      Component.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
      Component.AddParameter(pdtDateTime, CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
      Component.AddParameter(pdtDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)

      Dim Indicador As Integer
      Indicador = Component.Execute("BuscarIndicardorMultiplosVinculosEmVigenciaCruzada")

      If (Indicador = -1) Then
        bsShowMessage("Operação não permitida. Existe mais de um comprovante com vigência cruzada e indicador de múltiplos vinculos diferentes.", "E")
        CanContinue = False
      ElseIf (Indicador > 0 And Indicador <> CurrentQuery.FieldByName("INDMV").AsInteger) Then
        CurrentQuery.FieldByName("INDMV").AsInteger = Indicador
        bsShowMessage("Indicador de múltiplos vínculos alterado pois a vigência está cruzando com outro comprovante existente.", "I")
      End If

      DestroyComponent
    End If
End Sub

Public Function ExisteVinculoESocial(HandleComprovanteInss As Integer) As Boolean
  CreateComponent

  Component.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  ExisteVinculoESocial = Component.Execute("ExisteVinculoESocial")

  DestroyComponent
End Function

Public Sub CreateComponent()
  Set Component = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.Impostos.SamPrestadorINSSComprovanteBLL, Benner.Saude.Prestadores.Business")
  Component.ClearParameters
End Sub

Public Sub DestroyComponent()
  Set Component = Nothing
End Sub
