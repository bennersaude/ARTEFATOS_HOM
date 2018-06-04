'HASH: C4512D0262994CF7003B7F05D35CE4E2
'Macro: SAM_PROPONENTE

' Fábio -SMS 4570 -16/10/2001 -Considerar os 4 padrões de códigos diferentes para o prestador.
'         1: CNPJ/CPF,Registro
'         2: Registro,CNPJ/CPF
'         3: Código automático
'         4: Código manual

'#Uses "*VerificaEmail"
'#Uses "*CheckCPFCNPJ"
'#Uses "*bsShowMessage"

'Tipo do Código do prestador
Dim CODIGOPRESTADOR As Long
Dim EXIGIRCNPJCPF As String
Dim LIVREESCOLHACATEGORIA As Long
Dim MascaraInteiro As String

Option Explicit
Dim NaoFaturarGuiasAnterior As String

Public Sub BOTAOENVIARCONVITE_OnClick()

  If (CurrentQuery.State = 2 Or CurrentQuery.State = 3)  Then
    bsShowMessage("Ação não permitida. A fase está em edição.","I")
	Exit Sub
  End If

  On Error GoTo erro

	  Dim TvFormEnviarEmailBLL As CSBusinessComponent
	  Dim mensagemErro As String

	  Set TvFormEnviarEmailBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.TvFormEnviarEmailBLL, Benner.Saude.Prestadores.Business")
	  TvFormEnviarEmailBLL.AddParameter(pdtInteger, 3)
	  TvFormEnviarEmailBLL.AddParameter(pdtString, CurrentQuery.FieldByName("PROPONENTE").AsString)
	  TvFormEnviarEmailBLL.AddParameter(pdtInteger, -1)
	  TvFormEnviarEmailBLL.AddParameter(pdtInteger, -1)
	  TvFormEnviarEmailBLL.AddParameter(pdtInteger, 0)
	  TvFormEnviarEmailBLL.AddParameter(pdtAutomatic, False)
	  TvFormEnviarEmailBLL.Execute("PreencherFormularioEnvioEmail")
	  Set TvFormEnviarEmailBLL = Nothing
	  Exit Sub

  erro:
  	bsShowMessage(Err.Description, "E")
  	Set TvFormEnviarEmailBLL = Nothing
  Exit Sub

End Sub

Public Sub CEP_OnPopup(ShowPopup As Boolean)
  Dim vHandle As String
  Dim Interface As Object
  ShowPopup = False
  Set Interface = CreateBennerObject("ProcuraCEP.Rotinas")
  Interface.Exec(CurrentSystem, vHandle)

  If vHandle <>"" Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO,COMPLEMENTO   ")
    SQL.Add("  FROM LOGRADOUROS      ")
    SQL.Add(" WHERE CEP = :HANDLE ")
    SQL.ParamByName("HANDLE").Value = vHandle
    SQL.Active = True

    CurrentQuery.Edit
    CurrentQuery.FieldByName("CEP").Value = SQL.FieldByName("CEP").AsString
    CurrentQuery.FieldByName("ESTADO").Value = SQL.FieldByName("ESTADO").AsString
    CurrentQuery.FieldByName("MUNICIPIO").Value = SQL.FieldByName("MUNICIPIO").AsString
    CurrentQuery.FieldByName("BAIRRO").Value = SQL.FieldByName("BAIRRO").AsString
    CurrentQuery.FieldByName("LOGRADOURO").Value = SQL.FieldByName("LOGRADOURO").AsString
    CurrentQuery.FieldByName("LOGRADOUROCOMPLEMENTO").Value = SQL.FieldByName("COMPLEMENTO").AsString

  End If

  Set Interface = Nothing

End Sub

Public Sub FISICAJURIDICA_OnChange()
  Select Case Len(CurrentQuery.FieldByName("PROPONENTE").AsString)
    Case 0
      'se estiver vazio o PROPONENTE pode colocar como física ou jurídica
    Case 11
      CurrentQuery.FieldByName("FISICAJURIDICA").Value = 1
    Case 14
      CurrentQuery.FieldByName("FISICAJURIDICA").Value = 2
    Case Else
      CurrentQuery.FieldByName("FISICAJURIDICA").Clear
  End Select
  If CurrentQuery.State <>1 Then CurrentQuery.UpdateRecord
End Sub

Public Function LimpaCPFCNPJ(pCPFCNPJ As String) As String
  Dim viPos As Integer
  viPos = InStr(pCPFCNPJ, ".")
  While viPos>0
    pCPFCNPJ = Left(pCPFCNPJ, viPos -1) + Right(pCPFCNPJ, Len(pCPFCNPJ) - viPos)
    viPos = InStr(viPos, pCPFCNPJ, ".")
  Wend
  viPos = InStr(pCPFCNPJ, "-")
  While viPos>0
    pCPFCNPJ = Left(pCPFCNPJ, viPos -1) + Right(pCPFCNPJ, Len(pCPFCNPJ) - viPos)
    viPos = InStr(viPos, pCPFCNPJ, "-")
  Wend
  viPos = InStr(pCPFCNPJ, "/")
  While viPos>0
    pCPFCNPJ = Left(pCPFCNPJ, viPos -1) + Right(pCPFCNPJ, Len(pCPFCNPJ) - viPos)
    viPos = InStr(viPos, pCPFCNPJ, "/")
  Wend
  LimpaCPFCNPJ = pCPFCNPJ
End Function

Public Sub INDICADOPOR_OnChange()
  CurrentQuery.UpdateRecord
  CarregaIndicado(False)
End Sub

Public Sub PROPONENTE_Onexit()
  If CurrentQuery.State <>1 Then
    Dim vsCPFCNPJ As String
    vsCPFCNPJ = LimpaCPFCNPJ(CurrentQuery.FieldByName("PROPONENTE").AsString)
    Select Case Len(vsCPFCNPJ)
      Case 0
        If Len(CurrentQuery.FieldByName("PROPONENTE").AsString) = 0 Then
          CurrentQuery.FieldByName("PROPONENTE").Mask = ""
        Else
          bsShowMessage("Digite um CPF ou CNPJ válido !", "I")
        End If
        Exit Sub
      Case 11
        CurrentQuery.FieldByName("FISICAJURIDICA").Value = 1
        CurrentQuery.FieldByName("PROPONENTE").Value = vsCPFCNPJ
        CurrentQuery.FieldByName("PROPONENTE").Mask = "999\.999\.999\-99;0;_"
      Case 14
        CurrentQuery.FieldByName("FISICAJURIDICA").Value = 2
        CurrentQuery.FieldByName("PROPONENTE").Value = vsCPFCNPJ
        CurrentQuery.FieldByName("PROPONENTE").Mask = "99\.999\.999\/9999\-99;0;_"
      Case Else
        CurrentQuery.FieldByName("PROPONENTE").Mask = ""
        bsShowMessage("Digite um CPF ou CNPJ válido !", "I")
    End Select

    If CurrentQuery.State <>1 Then CurrentQuery.UpdateRecord

    Dim Pres As Object
    Set Pres = NewQuery
    Pres.Add("SELECT NOME FROM SAM_PRESTADOR WHERE PRESTADOR = :PROPONENTE")
    Pres.ParamByName("PROPONENTE").Value = CurrentQuery.FieldByName("PROPONENTE").AsString
    Pres.Active = True
    If Not Pres.EOF Then
      bsShowMessage("O Proponente " + Pres.FieldByName("NOME").AsString + " já cadastrado como Prestador", "I")
      Pres.Active = False
      Set Pres = Nothing
      CurrentQuery.FieldByName("PROPONENTE").Mask = ""
      Exit Sub
    End If
  End If
End Sub

Public Sub TABLE_AfterScroll()

	Dim SamPrestadorProcBLL As CSBusinessComponent
	Dim retorno As Boolean

	Set SamPrestadorProcBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcBLL, Benner.Saude.Prestadores.Business")
	SamPrestadorProcBLL.AddParameter(pdtString, "CREDENCIAMENTOAVANCADO")
	BOTAOENVIARCONVITE.Visible = SamPrestadorProcBLL.Execute("VerificarParametrosParaCredenciamentoAutomatico")
	Set SamPrestadorProcBLL = Nothing

	Select Case Len(CurrentQuery.FieldByName("PROPONENTE").AsString)
	Case 11
	  CurrentQuery.FieldByName("PROPONENTE").Mask = "999\.999\.999\-99;0;_"
	Case 14
	  CurrentQuery.FieldByName("PROPONENTE").Mask = "99\.999\.999\/9999\-99;0;_"
	Case Else
	  CurrentQuery.FieldByName("PROPONENTE").Mask = ""
	End Select

    CarregaIndicado(True)
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  '!!!VerificaCodigoPrestador(CanContinue)

  Dim Texto As String
  Dim CPFCGC As String
  Dim Pres As Object
  Dim Cont As Object
  Dim SQL As Object
  Dim vCont As Integer
  Dim vProtocolo As String



  'Checa Validade Do CGC ou CPF
  CPFCGC = CurrentQuery.FieldByName("PROPONENTE").AsString
  If Len(CPFCGC) = 11 Then
    If Not IsValidCPF(CPFCGC)Then
      bsShowMessage("CPF Inválido !", "E")
      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 2 Then '1-fisica 2-juridica
      CanContinue = False
      bsShowMessage("Para pessoa jurídica informe CNPJ !", "E")
      Exit Sub
    End If
  ElseIf Len(CPFCGC) = 14 Then
    If Not IsValidCGC(CPFCGC)Then
      bsShowMessage("CNPJ Inválido !", "E")
      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then '1-fisica 2-juridica
      CanContinue = False
      bsShowMessage("Para pessoa física informe CPF !", "E")
      Exit Sub

    End If
  Else
    bsShowMessage("CPF / CNPJ Inválido", "E")
    CanContinue = False
    Exit Sub
  End If


  ' Checa se ha algum prestador ja cadastrado
  Set Pres = NewQuery
  Pres.Add("SELECT NOME FROM SAM_PRESTADOR WHERE PRESTADOR = :PROPONENTE")
  Pres.ParamByName("PROPONENTE").Value = CurrentQuery.FieldByName("PROPONENTE").AsString
  Pres.Active = True
  If Not Pres.EOF Then
    bsShowMessage("O Proponente " + Pres.FieldByName("NOME").AsString + " já cadastrado como Prestador", "E")
    Pres.Active = False
    CanContinue = False
    Set Pres = Nothing
    Exit Sub
  End If

  'Checa se ja Existe um proponente ja cdastrado
  If CurrentQuery.State = 3 Then
    Dim Pres2 As Object
    Set Pres2 = NewQuery
    Pres2.Add("SELECT NOME FROM SAM_PROPONENTE WHERE PROPONENTE = :PROPONENTE AND SITUACAO <> 'I'")
    Pres2.ParamByName("PROPONENTE").Value = CurrentQuery.FieldByName("PROPONENTE").AsString
    Pres2.Active = True
    If Not Pres2.EOF Then
      CanContinue = False
      bsShowMessage("O Proponente " + Pres2.FieldByName("NOME").AsString + " já possui esse CPF/CNPJ", "E")
      Pres2.Active = False
      Set Pres2 = Nothing
      Exit Sub
    End If
  End If

  'Se situacao =c(contratado)exige data da contratacao
  If CurrentQuery.FieldByName("SITUACAO").AsString = "C" Then
    If CurrentQuery.FieldByName("DATACONTRATACAO").IsNull Then
      bsShowMessage("Data Contratação Obrigatória", "E")
      CanContinue = False
      DATACADASTRO.SetFocus
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 0 Then
    CanContinue = False
    bsShowMessage("Preencha Dados da Pessoa Física ou Jurídica!", "E")
    Exit Sub
  End If

  'limpar dados do tab tipoprestador
  Select Case CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger
    Case 1
      CurrentQuery.FieldByName("INSCRICAOESTADUAL").Clear
      CurrentQuery.FieldByName("INSCRICAOMUNICIPAL").Clear
      CurrentQuery.FieldByName("INSCRICAOINSS").Clear
      CurrentQuery.FieldByName("CORPOCLINICO").Value = "N"
      CurrentQuery.FieldByName("ASSOCIACAO").Value = "N"
      CurrentQuery.FieldByName("COOPERATIVA").Value = "N"
    Case 2
      CurrentQuery.FieldByName("DATANASCIMENTO").Clear
      CurrentQuery.FieldByName("ESTADOCIVIL").Clear
      CurrentQuery.FieldByName("SEXO").Clear
      CurrentQuery.FieldByName("RG").Clear
      CurrentQuery.FieldByName("ORGAOEMISSOR").Clear
      CurrentQuery.FieldByName("NATURALIDADE").Clear
      CurrentQuery.FieldByName("NACIONALIDADE").Clear
      CurrentQuery.FieldByName("CENTRALPAGER").Clear
      CurrentQuery.FieldByName("PAGER").Clear
  End Select

  'verifcação de preenchimento de datas
  If CanContinue = True Then
    If CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime >ServerNow Then
      CanContinue = False
      bsShowMessage("Data de inclusão não pode ser maior que a data atual!", "E")
      Exit Sub
    End If
    If CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime >ServerDate Then
      CanContinue = False
      bsShowMessage("Data de nascimento não pode ser maior que a data atual!", "E")
      Exit Sub
    End If
    If CurrentQuery.FieldByName("DATAINSCRICAOCR").AsDateTime >ServerDate Then
      CanContinue = False
      bsShowMessage("Data de inclusão no Conselho Regional não pode ser maior que a data atual!", "E")
      Exit Sub
    End If
    If CurrentQuery.FieldByName("DATACADASTRO").AsDateTime >ServerDate Then
      CanContinue = False
      bsShowMessage("Data de Cadastro não pode ser maior que a data atual!", "E")
      Exit Sub
    End If
  End If

  'Inscrição do INSS deve ser igual ao CNPJ se ambos forem informados
  If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 2 Then
    If CurrentQuery.FieldByName("INSCRICAOINSS").AsString <>"" Then
      If CurrentQuery.FieldByName("INSCRICAOINSS").AsString <>CurrentQuery.FieldByName("PROPONENTE").AsString Then
        bsShowMessage("Inscrição do INSS deve ser igual ao CNPJ.", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery
    qPermissao.Active = False
    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("FROM   Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add("WHERE  A.FILIAL = :FILIAL")
    qPermissao.Add("AND    A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True
    If qPermissao.FieldByName("ALTERAR").AsString <>"S" Then
      bsShowMessage( "Permissão negada! Usuário não pode alterar", "E")
      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If
    Set qPermissao = Nothing
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery

    qPermissao.Active = False
    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("FROM   Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add("WHERE  A.FILIAL = :FILIAL")
    qPermissao.Add("AND    A.USUARIO = :USUARIO")
    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True

    If qPermissao.FieldByName("INCLUIR").AsString <>"S" Then
      bsShowMessage("Permissão negada! Usuário não pode incluir", "E")
      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If

    Set qPermissao = Nothing
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If (VisibleMode) Then
    Dim qPermissao As Object
    Set qPermissao = NewQuery
    qPermissao.Active = False

    qPermissao.Add("SELECT A.ALTERAR, A.EXCLUIR, A.INCLUIR")
    qPermissao.Add("FROM   Z_GRUPOUSUARIOS_FILIAIS A")
    qPermissao.Add("WHERE  A.FILIAL = :FILIAL")
    qPermissao.Add("AND    A.USUARIO = :USUARIO")

    qPermissao.ParamByName("USUARIO").Value = CurrentUser
    qPermissao.ParamByName("FILIAL").Value = RecordHandleOfTable("FILIAIS")
    qPermissao.Active = True
    If qPermissao.FieldByName("EXCLUIR").AsString <>"S" Then
      bsShowMessage("Permissão negada! Usuário não pode excluir", "E")
      CanContinue = False
      Set qPermissao = Nothing
      Exit Sub
    End If

    Set qPermissao = Nothing
  End If
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("filial").Value = RecordHandleOfTable("FILIAIS")
End Sub

Public Sub CarregaIndicado(AfterScroll As Boolean)
  Select Case(CurrentQuery.FieldByName("INDICADOPOR").AsString)
	Case "B"
      INDICADOPORPRESTADOR.Visible = False
	  INDICADOPORBENEFICIARIO.Visible = True
	  If Not AfterScroll Then
      	CurrentQuery.FieldByName("INDICADOPORPRESTADOR").Clear
      End If
	Case "P"
	  INDICADOPORBENEFICIARIO.Visible = False
	  INDICADOPORPRESTADOR.Visible = True
	  If Not AfterScroll Then
	    CurrentQuery.FieldByName("INDICADOPORBENEFICIARIO").Clear
	  End If
	Case "O"
	  INDICADOPORBENEFICIARIO.Visible = False
	  INDICADOPORPRESTADOR.Visible = False
	  If Not AfterScroll Then
	    CurrentQuery.FieldByName("INDICADOPORPRESTADOR").Clear
	    CurrentQuery.FieldByName("INDICADOPORBENEFICIARIO").Clear
	  End If
  End Select
End Sub
