'HASH: EAAE3551DFF2529A3FB8E2576BDB0C78

'MACRO SFN_PESSOA_INSSCOMPROVANTE
'#Uses "*bsShowMessage"


Option Explicit

Public Sub TABLE_AfterDelete()
  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SFN_PSSOA_INSSCOMPROVANTE")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "Z")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_AfterPost()
  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SFN_PESSOA_INSSCOMPROVANTE")
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

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
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

  If ((Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull) And  (CurrentQuery.FieldByName("COMPETENCIA").AsDateTime > CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)) Then
	  bsShowMessage("A Competência não pode ser superior a Competência Final.", "E")
	  CanContinue = False
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
End Sub
