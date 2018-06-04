'HASH: 91FCE055A08EEE9ACBBB497F63B31BDA
'Macro: SAM_TGE

Dim vgNovaEstrutura As String
Dim EstadoAnterior As String
Dim DescricaoOld As String
Dim ClasseEventoOld As Long

'#Uses "*Modulo11"
'#Uses "*bsShowMessage"


Public Sub BOTAOALERTAGERALEVENTO_OnClick()
  Dim GeraAlertaDLL As Object
  Set GeraAlertaDLL = CreateBennerObject("TGE.Rotinas")
  GeraAlertaDLL.GeraAlertaEvento(CurrentSystem, CurrentQuery.FieldByName("ESTRUTURA").AsString)
  Set GeraAlertaDLL = Nothing
End Sub

Public Sub BOTAODUPLICARALTERACOES_OnClick()

  If CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "N" Then
	bsShowMessage("O evento a ser duplicado não é último nível!", "I")
	Exit Sub
  End If

  Dim DuplicaAlteracoesDLL As Object
  Set DuplicaAlteracoesDLL = CreateBennerObject("SamReplicTGE.Replica")
  DuplicaAlteracoesDLL.Duplicar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 256)
  Set DuplicaAlteracoesDLL = Nothing
End Sub

Public Sub BOTAODUPLICAREQUIVALENTES_OnClick()
  Dim vDuplicarEquivalentesDLL As Object
  Set vDuplicarEquivalentesDLL = CreateBennerObject("SamReplicTGE.EventoEquivalente")
  vDuplicarEquivalentesDLL.Duplicar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 256)
  Set vDuplicarEquivalentesDLL = Nothing
End Sub

Public Sub BOTAOPADRONIZAREQUIVALENTES_OnClick()
  Dim vPadronizarEquivalentesDLL As Object
  Set vPadronizarEquivalentesDLL = CreateBennerObject("SamReplicTGE.EventoEquivalente")
  vPadronizarEquivalentesDLL.Padronizar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set vPadronizarEquivalentesDLL = Nothing
End Sub

Public Sub BOTAOREPLICARALTERACOES_OnClick()
  Dim ReplicaAlteracoesDLL As Object
  Set ReplicaAlteracoesDLL = CreateBennerObject("SamReplicTGE.Replica")
  ReplicaAlteracoesDLL.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set ReplicaAlteracoesDLL = Nothing
End Sub

Public Sub CBHPMTABELA_OnChange()
  CurrentQuery.FieldByName("CBHPMESTRUTURA").Clear
  CurrentQuery.FieldByName("CBHPMESTRUTURANUMERICA").Clear
End Sub

Public Sub CBHPMTABELA_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long

  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")

  vHandle = Interface.Exec(CurrentSystem, "SAM_CBHPM", "ESTRUTURA|DESCRICAO", 1, "Estrutura|Descrição", " ULTIMONIVEL = 'S' ", "Tabela CBHPM", True, "")

  If vHandle > 0 Then
    CurrentQuery.FieldByName("CBHPMTABELA").Value = vHandle
  End If

  Set Interface = Nothing
End Sub

Public Sub DESCRICAO_OnExit()
  If CurrentQuery.FieldByName("DESCRICAOABREVIADA").AsString = "" And (CurrentQuery.State = 2 Or CurrentQuery.State = 3) Then                                                                                                        'diasaqui
    CurrentQuery.FieldByName("DESCRICAOABREVIADA").AsString = CurrentQuery.FieldByName("DESCRICAO").AsString
  End If
End Sub

Public Sub ESTRUTURA_OnExit()
  If (CurrentQuery.State = 3) And (CurrentQuery.FieldByName("ESTRUTURAAUXILIAR").AsString <> CurrentQuery.FieldByName("ESTRUTURA").AsString) Then
    CurrentQuery.FieldByName("ESTRUTURAAUXILIAR").Value = CurrentQuery.FieldByName("ESTRUTURA").AsString
  End If
End Sub

Public Sub TABLE_AfterPost()

  If vgNovaEstrutura <> "" Then
    Dim qUpdate As BPesquisa
    Dim vIndice As Long

    Set qUpdate = NewQuery

    Dim vEstruturaSemMascara As String
    vEstruturaSemMascara = ""
    For vIndice = 1 To Len(vgNovaEstrutura)
      If InStr("0123456789", Mid(vgNovaEstrutura, vIndice, 1)) > 0 Then
        vEstruturaSemMascara = vEstruturaSemMascara + Mid(vgNovaEstrutura, vIndice, 1)
      End If
    Next vIndice

    qUpdate.Clear
    qUpdate.Add("UPDATE SAM_TGE                               ")
    qUpdate.Add("   SET ESTRUTURA = :ESTRUTURA,               ")
    qUpdate.Add("       ESTRUTURAAUXILIAR = :ESTRUTURA,       ")
    qUpdate.Add("       ESTRUTURANUMERICA = :ESTRUTURANUMERICA")
    qUpdate.Add(" WHERE HANDLE = :HANDLE                      ")
    qUpdate.ParamByName("ESTRUTURA").AsString = vgNovaEstrutura
    qUpdate.ParamByName("ESTRUTURANUMERICA").AsInteger = Val(vEstruturaSemMascara)
    qUpdate.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qUpdate.ExecSQL
    vgNovaEstrutura = ""
    Set qUpdate = Nothing
  End If

  If (CurrentQuery.FieldByName("INATIVO").AsString <> EstadoAnterior) And _
     (EstadoAnterior <> "") And _
     CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "N" Then

      If CurrentQuery.FieldByName("INATIVO").AsString = "N" Then
        If WebMode Then
          TrocaEstadoFilhos(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("INATIVO").AsString)
          bsShowMessage("Os registros filhos foram ativados!", "I")
        Else
          If bsShowMessage("Deseja ativar os filhos?", "Q") = vbYes Then
            TrocaEstadoFilhos(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("INATIVO").AsString)
          End If
        End If
      Else
        TrocaEstadoFilhos(CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("INATIVO").AsString)
      End If
  End If

  Dim callEntity As CSEntityCall

  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamTge, Benner.Saude.Entidades", "EnviaProcedimentoParaSincronizacaoIntegracaoHospitalar")
  callEntity.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  callEntity.AddParameter(pdtString, DescricaoOld)
  callEntity.AddParameter(pdtString, CurrentQuery.FieldByName("DESCRICAO").AsString)
  callEntity.AddParameter(pdtAutomatic, ClasseEventoOld)
  callEntity.AddParameter(pdtAutomatic, IIf(IsNull(CurrentQuery.FieldByName("CLASSEEVENTO").AsInteger),0,CurrentQuery.FieldByName("CLASSEEVENTO").AsInteger))

  callEntity.Execute

  Set callEntity = Nothing

End Sub

Public Sub TABLE_AfterScroll()

'Inicio SMS  174752 - Leandro Manso - 12/12/2011
 If CurrentQuery.FieldByName("TABTIPOEVENTO").AsInteger = 5 Then
	UNIDADE.Caption = "Apresentação:"
	FORNECEDORPESSOA.Caption = "Marca/Responsável:"
  Else
  	UNIDADE.Caption = "Unidade de Fração/Medida"
	FORNECEDORPESSOA.Caption = "Empresa/Laboratório Responsável"
  End If
'Fim SMS 174752 - Leandro Manso - 12/12/2011

  If (WebMode) Then
    If CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "N" Then
      INTERVENCAO.ReadOnly = True
    Else
      INTERVENCAO.ReadOnly = False
    End If
    CBHPMTABELA.WebLocalWhere = " A.ULTIMONIVEL = 'S' "
  Else
    If CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "N" Then
      INTERVENCAO.Visible = False
    Else
      INTERVENCAO.Visible = True
    End If
  End If

  EstadoAnterior = CurrentQuery.FieldByName("INATIVO").AsString

  Dim qMascaraTGE As BPesquisa
  Set qMascaraTGE = NewQuery

  qMascaraTGE.Add("SELECT MASCARA, NIVEIS")
  qMascaraTGE.Add("  FROM SAM_MASCARATGE")
  qMascaraTGE.Add(" WHERE HANDLE = :HMASCARATGE")

  If CurrentQuery.State = 1 Then

    qMascaraTGE.ParamByName("HMASCARATGE").AsInteger = CurrentQuery.FieldByName("MASCARATGE").AsInteger
    qMascaraTGE.Active = True

    DynamicMask(qMascaraTGE.FieldByName("MASCARA").AsString, qMascaraTGE.FieldByName("NIVEIS").AsInteger, "SAM_TGE")

  Else
    If (CurrentQuery.State = 3 And VisibleMode) Then
      qMascaraTGE.ParamByName("HMASCARATGE").AsInteger = RecordHandleOfTable("SAM_MASCARATGE")
      qMascaraTGE.Active = True

      DynamicMask(qMascaraTGE.FieldByName("MASCARA").AsString, qMascaraTGE.FieldByName("NIVEIS").AsInteger, "SAM_TGE")
    End If
  End If

  Set qMascaraTGE = Nothing



	Dim qAux As BPesquisa
	Set qAux = NewQuery

	qAux.Add("SELECT COUNT(*) QTD                       ")
	qAux.Add("  FROM SAM_TGE_GRAU TG                    ")
	qAux.Add("  JOIN SAM_GRAU G ON (G.HANDLE = TG.GRAU) ")
	qAux.Add(" WHERE EVENTO = :PEVENTO                  ")
	qAux.Add("   AND G.ORIGEMVALOR = '2'               ")

	qAux.ParamByName("PEVENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	qAux.Active  = True

	If qAux.FieldByName("QTD").AsInteger = 0 Then
	    QTDAUX.Text = "**** Evento não possui auxiliares cadastrados ****"
	Else
		If qAux.FieldByName("QTD").AsInteger = 1 Then
	    	QTDAUX.Text = "Evento possui " & Str(qAux.FieldByName("QTD").AsInteger) & " auxiliar nos graus válidos."
	    Else
	    	QTDAUX.Text = "Evento possui " & Str(qAux.FieldByName("QTD").AsInteger) & " auxiliares nos graus válidos."
	    End If
	End If

    qAux.Active = False
    qAux.Clear
    qAux.Add("SELECT TABREGRAPERCENTUALUCOPOR   ")
    qAux.Add("  FROM SAM_PARAMETROSPRESTADOR    ")
    qAux.Active = True

    If qAux.FieldByName("TABREGRAPERCENTUALUCOPOR").AsInteger = 1 Then
      APLICARPERCENTUALPAGTO.Visible = False
    Else
      APLICARPERCENTUALPAGTO.Visible = True
    End If

	Set qAux = Nothing

	DescricaoOld = IIf(IsNull(CurrentQuery.FieldByName("DESCRICAO").AsString),"",CurrentQuery.FieldByName("DESCRICAO").AsString)
    ClasseEventoOld = IIf(IsNull(CurrentQuery.FieldByName("CLASSEEVENTO").AsInteger),0,CurrentQuery.FieldByName("CLASSEEVENTO").AsInteger)

End Sub

Public Sub TABLE_AfterInsert()

  If (CurrentQuery.State = 2 Or CurrentQuery.State = 3) And _
     Not CurrentQuery.FieldByName("NIVELSUPERIOR").IsNull Then
    Dim SQL As BPesquisa
    Set SQL = NewQuery

    SQL.Add("SELECT TIPOSERVICO FROM SAM_TGE WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("NIVELSUPERIOR").AsInteger
    SQL.Active = True

    If SQL.FieldByName("TIPOSERVICO").IsNull Then
      bsShowMessage("Tipo de Serviço do nível superior está nulo", "I")
    Else
      CurrentQuery.FieldByName("TIPOSERVICO").Value = SQL.FieldByName("TIPOSERVICO").AsInteger
    End If
    Set SQL = Nothing
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vMensagemQuestionamento As String
  vMensagemQuestionamento = ""

  If (CurrentQuery.FieldByName("INATIVO").AsString <> EstadoAnterior) And _
     (EstadoAnterior <> "") Then ' alterou a situação de ativo para inativo ou inativo para ativo

    If (CurrentQuery.FieldByName("ULTIMONIVEL").AsString = "N") And _
       (CurrentQuery.FieldByName("INATIVO").AsString = "S") Then ' alterou de ativo para inativo e não é ultimo nível, portanto verificar se é isso mesmo que ele quer fazer

      If VisibleMode Then

        If bsShowMessage("Os eventos de níveis inferiores serão inativados! Deseja continuar", "Q") = vbNo  Then
          CanContinue = False
          Exit Sub
        End If
      Else
        vMensagemQuestionamento = "Os eventos de níveis inferiores serão inativados!"
      End If

    Else ' Evento estava inativo e agora quer deixá-lo como ativo checar "pai" para ver se pode ativar

      If CurrentQuery.FieldByName("NIVELSUPERIOR").AsInteger > 0 Then ' existe pai entao checar
        Dim qChecaSituacaoAtivo As Object
        Set qChecaSituacaoAtivo = NewQuery
        qChecaSituacaoAtivo.Clear
        qChecaSituacaoAtivo.Add("SELECT HANDLE, INATIVO FROM SAM_TGE WHERE HANDLE = :NIVELSUPERIOR")
        qChecaSituacaoAtivo.ParamByName("NIVELSUPERIOR").AsInteger = CurrentQuery.FieldByName("NIVELSUPERIOR").AsInteger
        qChecaSituacaoAtivo.Active = True
        If (qChecaSituacaoAtivo.FieldByName("HANDLE").AsInteger > 0) And _
           (qChecaSituacaoAtivo.FieldByName("INATIVO").AsString = "S") Then ' pai encontrado e está inativo
          bsShowMessage("O evento de nível superior está INATIVO. Deve-se ativá-lo antes de efetivar a operação!", "E")
          CanContinue = False
          Set qChecaSituacaoAtivo = Nothing
          Exit Sub
        End If
        Set qChecaSituacaoAtivo = Nothing

      End If

    End If



  End If
  CurrentQuery.FieldByName("ESTRUTURAAUXILIAR").AsString = CurrentQuery.FieldByName("ESTRUTURA").AsString


  If CurrentQuery.FieldByName("FATORCONTAGEM").Value = 0 Then
    CanContinue = False
    bsShowMessage("Fator de contagem deve ser maior que zero", "E")
    Exit Sub
  End If

  Dim qParam As BPesquisa
  Set qParam = NewQuery

  qParam.Active = False
  qParam.Clear
  qParam.Add("SELECT CALCCODPAGTOEVENTOCIRURGICO, ")
  qParam.Add("       EXIGIRPARECERPERICIA,        ")
  qParam.Add("       FORNECIMENTOMEDICAMENTO      ")
  qParam.Add("  FROM SAM_PARAMETROSATENDIMENTO    ")
  qParam.Active = True

    If (qParam.FieldByName("EXIGIRPARECERPERICIA").AsString) = "S" Then
      If (CurrentQuery.FieldByName ("REGIMEATENDIMENTO").AsInteger = 1) Then
        CanContinue = True
      ElseIf (CurrentQuery.FieldByName ("EXAMEPREOPERATORIO").AsString = "N" )Then
        If (VisibleMode) Then
          ' Julio - SMS: 71657 - Início
          ' Sms 148416 - 09/11/2010
          If bsShowMessage("Regime de atendimento diferente de ambulatorial, deseja salvar sem exigir perícia prévia ?", "Q") = vbYes Then
            CanContinue = True
          Else
            Set qParam = Nothing
            CanContinue = False
            Exit Sub
          End If
          ' Fim sms 148416
          ' Julio - SMS: 71657 - Fim
        Else
          If vMensagemQuestionamento <> "" Then
            vMensagemQuestionamento = vMensagemQuestionamento +  Chr(13) + Chr(10) + "Regime de atendimento diferente de ambulatorial, deseja salvar sem exigir perícia prévia ?"
          Else
            vMensagemQuestionamento = "Regime de atendimento diferente de ambulatorial, deseja salvar sem exigir perícia prévia ?"
          End If
        End If
      End If
    End If

  If vMensagemQuestionamento <> "" Then
    RequestConfirmation(vMensagemQuestionamento)
  End If


  Dim VmascaraTGE As String

  Dim qMascaraTGE As BPesquisa
  Set qMascaraTGE = NewQuery

  qMascaraTGE.Add("SELECT DESCRICAO, MASCARA, VALIDARDIGITOVERIFICADOR")
  qMascaraTGE.Add("  FROM SAM_MASCARATGE")
  qMascaraTGE.Add(" WHERE HANDLE = :HMASCARATGE")
  qMascaraTGE.ParamByName("HMASCARATGE").AsInteger = CurrentQuery.FieldByName("MASCARATGE").AsInteger
  qMascaraTGE.Active = True

  VmascaraTGE = qMascaraTGE.FieldByName("MASCARA").AsString

  If (NiveisEstrutura(CurrentQuery.FieldByName("ESTRUTURA").AsString) = NiveisEstrutura(VmascaraTGE)) And _
       (Len(CurrentQuery.FieldByName("ESTRUTURA").AsString) <> Len(VmascaraTGE)) Then
    CanContinue = False
    bsShowMessage("Para último nível todo o código do evento deve ser preenchido", "E")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("PREEXISTENCIA").AsString = "S" Or _
      CurrentQuery.FieldByName("PREEXISTENCIA").AsString = "P") Then
    If CurrentQuery.FieldByName("MESESPREEXISTENCIA").IsNull Then
      CanContinue = False
      bsShowMessage("Deve ser informado a quantidade de meses em que se observa casos de pré-existencia", "E")
      Exit Sub
    End If
  Else
    CurrentQuery.FieldByName("MESESPREEXISTENCIA").Value = Null
  End If

  'juliana alterado em 03/06/2002 para verificar alteração na tabela de niveis
  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
    Dim QueryNiveis As BPesquisa
    Set QueryNiveis = NewQuery
    QueryNiveis.Clear

    QueryNiveis.Add("SELECT N.NIVEL FROM SAM_NIVELAUTORIZACAO N WHERE N.HANDLE=:NIVEL")
    QueryNiveis.ParamByName("NIVEL").Value = CurrentQuery.FieldByName("TABNIVELAUTORIZACAO").AsInteger
    QueryNiveis.Active = False
    QueryNiveis.Active = True

    CurrentQuery.FieldByName("NIVELAUTORIZACAO").AsInteger = QueryNiveis.FieldByName("NIVEL").AsInteger
    Set QueryNiveis = Nothing
  End If

  If Not CurrentQuery.FieldByName("OCORRENCIAMAXIMA").IsNull Then
    If CurrentQuery.FieldByName("OCORRENCIAMAXIMA").AsInteger <= 0 And CurrentQuery.FieldByName("PRONTUARIO").AsString<>"S" Then 'Sem prontuário
      'Alterado por Garcia - Não faz sentido setar o prontuário se existe uma verificação no beforePost.
      'CurrentQuery.FieldByName("PRONTUARIO").Value="S"
      bsShowMessage("A ocorrencia máxima deve ser maior que zero, se o evento deve ser registrado no prontuário", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  Dim SQL As BPesquisa
  Set SQL = NewQuery
  Dim SQL3 As BPesquisa
  Set SQL3 = NewQuery

  If qMascaraTGE.FieldByName("VALIDARDIGITOVERIFICADOR").AsString = "S" Then
    ' Tratamento da Estrutura da SAM_TGE
    Dim vEstruturaSemMascara As String
    vEstruturaSemMascara = ""

    Dim vQtdeNumerosMascara As Integer
    vQtdeNumerosMascara = 0

    Dim vIndice As Integer

    For vIndice = 1 To Len(CurrentQuery.FieldByName("ESTRUTURA").AsString)
      If InStr("0123456789", Mid(CurrentQuery.FieldByName("ESTRUTURA").AsString, vIndice, 1)) > 0 Then
        vEstruturaSemMascara = vEstruturaSemMascara + _
                               Mid(CurrentQuery.FieldByName("ESTRUTURA").AsString, vIndice, 1)
      End If
    Next vIndice

    For vIndice = 1 To Len(VmascaraTGE)
      If InStr("0123456789", Mid(VmascaraTGE, vIndice, 1)) > 0 Then
        vQtdeNumerosMascara = vQtdeNumerosMascara + 1
      End If
    Next vIndice

    SQL.Add("SELECT * FROM Z_MASCARAs WHERE TABELA = (SELECT HANDLE FROM Z_TABELAS WHERE NOME = 'SAM_TGE')")
    SQL.Active = True

    If Not SQL.EOF Then
      Dim vDigitoVerificador As Integer
      vDigitoVerificador = Val(Modulo11(Mid(vEstruturaSemMascara, 1, Len(vEstruturaSemMascara) -1)))
      If Len(CurrentQuery.FieldByName("ESTRUTURA").AsString) = Len(VmascaraTGE) Then
        CurrentQuery.FieldByName("ULTIMONIVEL").Value = "S"
        'Avisar se o dígito verificador do evento estiver incorreto
        If Right(vEstruturaSemMascara, 1) <> vDigitoVerificador And _
        	 (qMascaraTGE.FieldByName("DESCRICAO").AsString <> "TUSS") Then
            		If bsShowMessage("Dígito verificador inválido! Dígito correto seria " + Str(vDigitoVerificador) + " deseja alterar? ", "Q" ) = vbYes Then
              		vgNovaEstrutura = Left (CurrentQuery.FieldByName("ESTRUTURA").AsString, Len(VmascaraTGE) -1) + Trim(Str(vDigitoVerificador))

                      Dim vEstruturaSemMascaraE As String
                      Dim vIndiceE As Long
                      vEstruturaSemMascaraE = ""
                      For vIndiceE = 1 To Len(vgNovaEstrutura)
                        If InStr("0123456789", Mid(vgNovaEstrutura, vIndiceE, 1)) > 0 Then
                          vEstruturaSemMascaraE = vEstruturaSemMascaraE + Mid(vgNovaEstrutura, vIndiceE, 1)
                        End If
                      Next vIndiceE

                      Dim qExiste As BPesquisa
                      Set qExiste = NewQuery

  					qExiste.Clear
  					qExiste.Add("SELECT COUNT(*) QTDE FROM SAM_TGE            ")
  					qExiste.Add(" WHERE ESTRUTURA = :ESTRUTURA                ")
  					qExiste.Add("   AND ESTRUTURAAUXILIAR = :ESTRUTURA        ")
  					qExiste.Add("   AND ESTRUTURANUMERICA = :ESTRUTURANUMERICA")
  					qExiste.ParamByName("ESTRUTURA").AsString = vgNovaEstrutura
  					qExiste.ParamByName("ESTRUTURANUMERICA").AsInteger = Val(vEstruturaSemMascaraE)
                      qExiste.Active = True

                      If qExiste.FieldByName("QTDE").AsInteger > 0 Then
                        CanContinue = False
                        Set qExiste = Nothing
                        bsShowMessage("Nova estrutura com dígito correto, já existe para outro evento!","E")
                        Exit Sub
                      End If
            		Else
            		  If VisibleMode Then
              		    CanContinue = False
              		    Exit Sub
              		  End If
            		End If
        End If
      Else
        vEstruturaSemMascara = Left((vEstruturaSemMascara + "0000000000"), _
                               vQtdeNumerosMascara -1)
        vEstruturaSemMascara = Trim(vEstruturaSemMascara) + Trim(Str(vDigitoVerificador))
        CurrentQuery.FieldByName("ULTIMONIVEL").Value = "N"
      End If
    End If
  Else ' não vai fazer a consistência quanto a DV mais popula o UltimoNivel de acordo com o definido na MáscaraTGE
    If Len(CurrentQuery.FieldByName("ESTRUTURA").AsString) = Len(VmascaraTGE) Then
      CurrentQuery.FieldByName("ULTIMONIVEL").Value = "S"
    End If
  End If
  SQL3.Active = False
  Set qMascaraTGE = Nothing

  SQL3.Clear
  SQL3.Add("SELECT ST.HANDLE ")
  SQL3.Add("  FROM SAM_TGESUBSTITUTO_ITEM STI ")
  SQL3.Add("  JOIN SAM_TGESUBSTITUTO      ST  ON ST.HANDLE = STI.TGESUBSTITUTO ")
  SQL3.Add(" WHERE STI.EVENTO = :EVENTO")
  SQL3.Add("    OR ST.EVENTO = :EVENTO")
  SQL3.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL3.Active = True

  If Not SQL3.EOF Then
    If CurrentQuery.FieldByName("INATIVO").AsString = "S" Then
      bsShowMessage("O evento não pode ser inativado pois faz parte da uma lista de substituição", "E")
      CanContinue = False
      SQL3.Active = False
      Set SQL3 = Nothing
      Exit Sub
    End If
  End If

  SQL3.Active = False
  Set SQL3 = Nothing

  CurrentQuery.FieldByName("ESTRUTURANUMERICA").Value = Replace(CurrentQuery.FieldByName("ESTRUTURA").AsString, ".", "")

  'Se o evento é cirúrgico não pode ser informado o codigo do pagamento
  If (qParam.FieldByName("CALCCODPAGTOEVENTOCIRURGICO").AsString = "S") And _
     (CurrentQuery.FieldByName("CIRURGICO").AsString = "S") And _
     Not (CurrentQuery.FieldByName("CODIGOPAGTO").IsNull) Then
    CanContinue = False
    bsShowMessage("Está marcado nos parâmetros gerais que o percentual de pagamento" + Chr(13) + _
                  "será calculado pelo sistema para eventos cirúrgicos." + Chr(13) + _
                  "Este campo deverá ser deixado em branco!", "E")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("OCORRENCIAMAXIMA").AsInteger) > 0 And _
      (CurrentQuery.FieldByName("PRONTUARIO").AsString) = "S" Then
    bsShowMessage("Registro no prontuário não pode ser:" + Chr(13) + _
                  "Sem prontuário" + Chr(13) + _
                  "quando a ocorrência máxima for maior que zero!", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim qEspec As BPesquisa
  Set qEspec = NewQuery

  qEspec.Clear
  qEspec.Add("SELECT TABTIPOEVENTO, ESTRUTURA")
  qEspec.Add("  FROM SAM_TGE")
  qEspec.Add(" WHERE HANDLE = :HANDLE")
  qEspec.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("NIVELSUPERIOR").AsInteger
  qEspec.Active = True

  'Retirada a validação devido a nova tabela TUSS (TISS 3.1) possuir mais de uma especificação para a mesma estrutura do evento. SMS 287795
  'If Not CurrentQuery.FieldByName("NIVELSUPERIOR").IsNull Then
  '  'Cleber - Odontológico  15/09/2003
  '  If CurrentQuery.FieldByName("TABTIPOEVENTO").AsInteger <> 4 _
  '     And qEspec.FieldByName("TABTIPOEVENTO").AsInteger <> CurrentQuery.FieldByName("TABTIPOEVENTO").AsInteger _
  '     And Len(qEspec.FieldByName("ESTRUTURA").AsString ) > 1 Then
  '    bsShowMessage("'Especificação do Evento' está diferente do nível superior!", "E")
  '    CanContinue = False
  '    Set qEspec = Nothing
  '    Exit Sub
  '  End If
  'End If

  Set qEspec = Nothing

  If Not CurrentQuery.FieldByName("CBHPMTABELA").IsNull Then
    Dim SQLCBHPM As BPesquisa
    Set SQLCBHPM = NewQuery

    SQLCBHPM.Add("SELECT * FROM SAM_CBHPM WHERE HANDLE = :HANDLE")
    SQLCBHPM.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CBHPMTABELA").Value
    SQLCBHPM.Active = True

    CurrentQuery.FieldByName("CBHPMESTRUTURA").Value = SQLCBHPM.FieldByName("ESTRUTURA").Value
    CurrentQuery.FieldByName("CBHPMESTRUTURANUMERICA").Value = SQLCBHPM.FieldByName("ESTRUTURANUMERICA").Value

    Set SQLCBHPM = Nothing
  End If

  If (CurrentQuery.State = 3) Then
      Dim sql2 As BPesquisa
	  Set sql2 = NewQuery
	  sql2.Clear
	  sql2.Add("SELECT HANDLE                ")
	  sql2.Add("  FROM SAM_TGE               ")
	  sql2.Add(" WHERE ESTRUTURA = :ESTRUTURA")
	  sql2.Add("   AND MASCARATGE = :MASCARATGE")
	  sql2.ParamByName("ESTRUTURA").AsString = CurrentQuery.FieldByName("ESTRUTURA").AsString
	  sql2.ParamByName("MASCARATGE").AsInteger = CurrentQuery.FieldByName("MASCARATGE").AsInteger
	  sql2.Active = True

	  If (sql2.FieldByName("HANDLE").AsInteger > 0) Then
     	bsShowMessage(" Estrutura já cadastrada para esta máscara! ", "E")
      	CanContinue = False
      	Set sql2 = Nothing
      	Exit Sub
	  End If

	  Set sql2 = Nothing
  End If

  Dim qEventoAntes As BPesquisa
  Set qEventoAntes = NewQuery

  qEventoAntes.Clear
  qEventoAntes.Add("SELECT TABTIPOEVENTO")
  qEventoAntes.Add("  FROM SAM_TGE")
  qEventoAntes.Add(" WHERE HANDLE = :HANDLE")
  qEventoAntes.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qEventoAntes.Active = True

  If (qParam.FieldByName("FORNECIMENTOMEDICAMENTO").AsString <> "N" And _
      qEventoAntes.FieldByName("TABTIPOEVENTO").AsInteger <> CurrentQuery.FieldByName("TABTIPOEVENTO").AsInteger And _
      CurrentQuery.FieldByName("TABTIPOEVENTO").AsInteger = 4) Then

    Dim qEvento As BPesquisa
    Set qEvento = NewQuery

    qEvento.Clear
    qEvento.Add("SELECT COUNT(*) REGISTROS                                    ")
    qEvento.Add("  FROM SAM_TGE E                                             ")
    qEvento.Add("  JOIN SAM_TGE_COMPLEMENTAR C ON C.EVENTO = E.HANDLE         ")
    qEvento.Add(" WHERE (C.EVENTO = :HEVENTO OR C.EVENTOAGERAR = :HEVENTO)    ")
    qEvento.Add("   AND C.EVENTOAGERAR <> C.EVENTO                            ")
    qEvento.ParamByName("HEVENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qEvento.Active = True

    If qEvento.FieldByName("REGISTROS").AsInteger > 0 Then
      CanContinue = False
      Set qEvento = Nothing
      Set qEventoAntes = Nothing
      bsShowMessage("Não é possível alterar a 'Especificação do evento' para 'Medicamento'!"+ Chr(13) + _
                    "Pois, evento do tipo 'Medicamento' não possibilita 'Evento Complementar' diferente dele mesmo, quando há 'Cadastro de Fornecimento de Medicamento'!","E")
      Exit Sub
    End If

    Set qEvento = Nothing
  End If

  Set qEventoAntes = Nothing
  Set qParam = Nothing

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT TABINTEGRACAOORIZON FROM EMPRESAS WHERE HANDLE = :EMPRESA")
  SQL.ParamByName("EMPRESA").AsInteger = CurrentCompany
  SQL.Active = True

  If SQL.FieldByName("TABINTEGRACAOORIZON").AsString = "2" Then
    If CurrentQuery.FieldByName("EXAMEPREOPERATORIO").AsString = "S" And CurrentQuery.FieldByName("NUMEROBANDEJAAUDITORIA").IsNull  Then
      bsShowMessage("O campo ""Número Bandeja Auditoria"" deve ser preenchido", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If

  Set SQL = Nothing

End Sub


'Public Sub AtualizaEstruturaNumerica

'  Dim sValor As String
'  Dim i As Long'

'  sValor = ""

'  For i = 1 To Len(CurrentQuery.FieldByName("ESTRUTURA").AsString)
'    If InStr("0123456789", Mid(CurrentQuery.FieldByName("ESTRUTURA").AsString, i, 1)) > 0 Then
'      sValor = sValor + Mid(CurrentQuery.FieldByName("ESTRUTURA").AsString, i, 1)
'    End If
'  Next i

'  CurrentQuery.FieldByName("ESTRUTURANUMERICA").Value = Val(sValor)

'End Sub

'Public Sub AtualizaEstruturaNumericax

'  Dim SQL As Object
'  Dim sValor As String
'  Dim i As Long'

'  Set SQL = NewQuery

'  SQL.Add("Select * from sam_tge")
'  SQL.RequestLive = True
'  SQL.Active = True

'  While Not SQL.EOF

'    sValor = ""

'    For i = 1 To Len(SQL.FieldByName("ESTRUTURA").AsString)
'      If InStr("0123456789", Mid(SQL.FieldByName("ESTRUTURA").AsString, i, 1)) > 0 Then
'        sValor = sValor + Mid(SQL.FieldByName("ESTRUTURA").AsString, i, 1)
'      End If
'    Next i

'    SQL.Edit
'    SQL.FieldByName("ESTRUTURANUMERICA").Value = Val(sValor)
'    SQL.Post
'    SQL.Next

'  Wend

'End Sub

Public Function TrocaEstadoFilhos(pHandle As Long, pEstado As String)

Dim UP As BPesquisa
Set UP = NewQuery
Dim Consulta As BPesquisa
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT INATIVO, ULTIMONIVEL FROM SAM_TGE WHERE HANDLE = :PHANDLE")
Consulta.ParamByName("PHANDLE").AsInteger = pHandle
Consulta.Active = True

If Consulta.FieldByName("INATIVO").AsString <> pEstado Then
  UP.Clear
  UP.Add("UPDATE SAM_TGE SET INATIVO = :ESTADO WHERE HANDLE = :PHANDLE")
  UP.ParamByName("PHANDLE").AsInteger = pHandle
  UP.ParamByName("ESTADO").AsString = pEstado
  UP.ExecSQL
End If

If Consulta.FieldByName("ULTIMONIVEL").AsString <> "S" Then
  Consulta.Clear
  Consulta.Add("SELECT HANDLE FROM SAM_TGE WHERE NIVELSUPERIOR = :PHANDLE")
  Consulta.ParamByName("PHANDLE").AsInteger = pHandle
  Consulta.Active = True

  While Not Consulta.EOF
    TrocaEstadoFilhos(Consulta.FieldByName("HANDLE").AsInteger, pEstado)
    Consulta.Next
  Wend
End If

Set UP = Nothing
Set Consulta = Nothing

End Function

Public Sub TABLE_UpdateRequired()
  Dim qMascaraTGE As BPesquisa
  Set qMascaraTGE = NewQuery

  qMascaraTGE.Add("SELECT MASCARA, NIVEIS")
  qMascaraTGE.Add("  FROM SAM_MASCARATGE")
  qMascaraTGE.Add(" WHERE HANDLE = :HMASCARATGE")
  qMascaraTGE.ParamByName("HMASCARATGE").AsInteger = CurrentQuery.FieldByName("MASCARATGE").AsInteger
  qMascaraTGE.Active = True

  DynamicMask(qMascaraTGE.FieldByName("MASCARA").AsString, qMascaraTGE.FieldByName("NIVEIS").AsInteger, "SAM_TGE")

  Set qMascaraTGE = Nothing
  CurrentQuery.FieldByName("ESTRUTURAAUXILIAR").AsString = CurrentQuery.FieldByName("ESTRUTURA").AsString

End Sub
Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim qDel As Object
	Set qDel = NewQuery
	qDel.Add("DELETE FROM SAM_TGE_EQUIVALENTELOG")
	qDel.Add(" WHERE EVENTO = :EVENTO")
	qDel.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qDel.ExecSQL
    Set qDel = Nothing
End Sub

Public Sub TABTIPOEVENTO_OnChange()
'Inicio SMS  174752 - Leandro Manso - 12/12/2011
If TABTIPOEVENTO.PageIndex = 4 Then
	UNIDADE.Caption = "Apresentação:"
	FORNECEDORPESSOA.Caption = "Marca/Responsável:"
  Else
  	UNIDADE.Caption = "Unidade de Fração/Medida"
	FORNECEDORPESSOA.Caption = "Empresa/Laboratório Responsável"
  End If
'Fim SMS 174752 - Leandro Manso - 12/12/2011
End Sub

Public Sub TUSSEVENTO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long

  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")

  vHandle = Interface.Exec(CurrentSystem, _
  						   "SAM_TGE", _
  						   "ESTRUTURA|ESTRUTURANUMERICA|DESCRICAO", _
  						   3, _
  						   "Estrutura|Estrutura Númerica|Descrição", _
  						   "", _
  						   "Tabela", _
  						   False, "")

  If vHandle > 0 Then
    CurrentQuery.FieldByName("TUSSEVENTO").Value = vHandle
  End If

  Set Interface = Nothing
End Sub
