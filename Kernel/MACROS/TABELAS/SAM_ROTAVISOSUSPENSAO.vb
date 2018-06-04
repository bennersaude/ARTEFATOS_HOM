'HASH: D815387474C291D8B08B1D9430D89E84
'Macro: SAM_ROTAVISOSUSPENSAO

'JULIANO -22/08/2001

'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.FieldByName("processado").Value = "S" Then

    If Not InTransaction Then StartTransaction

    Dim PROCESSA As Object
    Dim DEST As Object
    Set PROCESSA = NewQuery
    Set DEST = NewQuery
    Dim APAGAATRASO As Object
    Set APAGAATRASO = NewQuery

    If MsgBox("Todos os dados já cadastrados serão apagados!" + Chr(13) + "Deseja Continuar?", 4) = vbYes Then


      APAGAATRASO.Active = False
      APAGAATRASO.Clear
      APAGAATRASO.Add("DELETE FROM SAM_ROTAVISOSUSPENSAO_ATRASO                   ")
      APAGAATRASO.Add(" WHERE EXISTS (Select A.HANDLE                             ")
      APAGAATRASO.Add("                    FROM SAM_ROTAVISOSUSPENSAO_ATRASO A,   ")
      APAGAATRASO.Add("                         SAM_ROTAVISOSUSPENSAO_DEST D,     ")
      APAGAATRASO.Add("                         SAM_ROTAVISOSUSPENSAO_CONTRATO C, ")
      APAGAATRASO.Add("                         SAM_ROTAVISOSUSPENSAO S           ")
      APAGAATRASO.Add("                   WHERE A.ROTINADEST = D.HANDLE           ")
      APAGAATRASO.Add("                     AND D.ROTINACONTRATO = C.HANDLE       ")
      APAGAATRASO.Add("                     And C.ROTINA = S.HANDLE               ")
      APAGAATRASO.Add("                     And A.HANDLE = SAM_ROTAVISOSUSPENSAO_ATRASO.HANDLE")
      APAGAATRASO.Add("                     And S.HANDLE = :HANDLEROTINA)         ")
      APAGAATRASO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      APAGAATRASO.ExecSQL

      DEST.Active = False
      DEST.Clear
      DEST.Add("DELETE FROM SAM_ROTAVISOSUSPENSAO_DEST                            ")
      DEST.Add(" WHERE EXISTS (SELECT ROTDEST.HANDLE                              ")
      DEST.Add("                    FROM SAM_ROTAVISOSUSPENSAO ROT,               ")
      DEST.Add("                         SAM_ROTAVISOSUSPENSAO_CONTRATO ROTCONT,  ")
      DEST.Add("                         SAM_ROTAVISOSUSPENSAO_DEST ROTDEST       ")
      DEST.Add("                   WHERE ROT.HANDLE = ROTCONT.ROTINA              ")
      DEST.Add("                     AND ROTCONT.HANDLE = ROTDEST.ROTINACONTRATO  ")
      DEST.Add("                     And ROTDEST.HANDLE = SAM_ROTAVISOSUSPENSAO_DEST.HANDLE")
      DEST.Add("                     AND ROT.HANDLE = :HANDLEROTINA)              ")
      DEST.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      DEST.ExecSQL

      PROCESSA.Active = False
      PROCESSA.Clear
      PROCESSA.Add("UPDATE SAM_ROTAVISOSUSPENSAO")
      PROCESSA.Add("   Set PROCESSADO = 'N'")
      PROCESSA.Add(" WHERE HANDLE = :HANDLEROTINA")
      PROCESSA.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      PROCESSA.ExecSQL

      WriteAudit("C", HandleOfTable("SAM_ROTAVISOSUSPENSAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Aviso de Suspensão - Cancelamento")

      RefreshNodesWithTable("SAM_ROTAVISOSUSPENSAO")

      If InTransaction Then Commit

    Else

      CurrentQuery.Cancel

    End If


    Set PROCESSA = Nothing
    Set DEST = Nothing
    Set APAGAATRASO = Nothing
  End If
End Sub

Public Sub BOTAOCRIARCONTRATO_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("O registro está em edição! Por favor confirme ou cancele as alterações")
    Exit Sub
  End If

  Dim CONT As Object
  Set CONT = CreateBennerObject("SamAvisoSuspensao.GerarAviso")
  CONT.GeraContrato(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set CONT = Nothing
End Sub

Public Sub BOTAOIMPRESSAO_OnClick()
  Dim vBenefConcatena As String
  Dim vPesConcatena As String
  Dim vBeneficiario As Long
  Dim vPessoa As Long
  Dim vRelatorio As String
  Dim vDestBen As Long
  Dim vDestPes As Long
  Dim vOcorrencias As String
  Dim vGravaOcorrencias As Boolean
  Dim BENE As Object
  Set BENE = NewQuery
  Dim PES As Object
  Set PES = NewQuery
  Dim QueryHandleRelatorio As Object
  Set QueryHandleRelatorio = NewQuery

  'Busca todos os beneficiários que foram gerados avisos,já com o relatório definido no contrato
  Dim DESTBEN As Object
  Set DESTBEN = NewQuery
  DESTBEN.Add("SELECT SD.HANDLE DEST,                       ")
  DESTBEN.Add("       SD.BENEFICIARIO,                      ")
  DESTBEN.Add("       C.RELATORIOAVISOSUSPENSAO,                     ")
  DESTBEN.Add("       C.CONTRATO,                           ")
  DESTBEN.Add("       SD.ROTINACONTRATO                     ")
  DESTBEN.Add("  FROM SAM_ROTAVISOSUSPENSAO_DEST SD,        ")
  DESTBEN.Add("       SAM_ROTAVISOSUSPENSAO_CONTRATO SC,    ")
  DESTBEN.Add("       SAM_ROTAVISOSUSPENSAO S,              ")
  DESTBEN.Add("       SAM_CONTRATO C                        ")
  DESTBEN.Add(" WHERE S.HANDLE = :HANDLEROTINA              ")
  DESTBEN.Add("   And SC.ROTINA = S.HANDLE                  ")
  DESTBEN.Add("   And SD.ROTINACONTRATO = SC.HANDLE         ")
  DESTBEN.Add("   And SD.CORRESPONDENCIA IS NULL            ")
  DESTBEN.Add("   And SD.BENEFICIARIO IS NOT NULL           ")
  DESTBEN.Add("   And SD.CONTRATO = C.HANDLE                ")
  DESTBEN.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DESTBEN.Active = True
  Dim DESTPES As Object
  Set DESTPES = NewQuery
  DESTPES.Add("SELECT SD.HANDLE DEST,                       ")
  DESTPES.Add("       SD.PESSOA,                            ")
  DESTPES.Add("       C.RELATORIOAVISOSUSPENSAO,                     ")
  DESTPES.Add("       C.CONTRATO,                           ")
  DESTPES.Add("       SD.ROTINACONTRATO                     ")
  DESTPES.Add("  FROM SAM_ROTAVISOSUSPENSAO_DEST SD,        ")
  DESTPES.Add("       SAM_ROTAVISOSUSPENSAO_CONTRATO SC,    ")
  DESTPES.Add("       SAM_ROTAVISOSUSPENSAO S,              ")
  DESTPES.Add("       SAM_CONTRATO C                        ")
  DESTPES.Add(" WHERE S.HANDLE = :HANDLEROTINA              ")
  DESTPES.Add("   And SC.ROTINA = S.HANDLE                  ")
  DESTPES.Add("   And SD.ROTINACONTRATO = SC.HANDLE         ")
  DESTPES.Add("   And SD.CORRESPONDENCIA Is Null            ")
  DESTPES.Add("   And SD.PESSOA Is Not Null                 ")
  DESTPES.Add("   And SD.CONTRATO = C.HANDLE                ")
  DESTPES.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DESTPES.Active = True
  If(DESTPES.EOF)And(DESTBEN.EOF)Then
  MsgBox("Não foi gerado nenhum aviso para esta rotina!")
  Exit Sub
End If
vBenefConcatena = ""
vPesConcatena = ""
vGravaOcorrencias = False
vOcorrencias = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + _
               "----------------------------------------------------------------------------------------------------- " + Str(ServerDate)
While Not DESTBEN.EOF
  vBeneficiario = DESTBEN.FieldByName("BENEFICIARIO").AsInteger
  vRelatorio = DESTBEN.FieldByName("RELATORIOAVISOSUSPENSAO").AsString
  vDestBen = DESTBEN.FieldByName("ROTINACONTRATO").AsInteger
  If vBeneficiario >0 Then
    If vBenefConcatena <>"" Then
      vBenefConcatena = vBenefConcatena + ","
    End If
    vBenefConcatena = vBenefConcatena + Str(vBeneficiario)
  End If
  DESTBEN.Next
Wend
While Not DESTPES.EOF
  vPessoa = DESTPES.FieldByName("PESSOA").AsInteger
  vRelatorio = DESTPES.FieldByName("RELATORIOAVISOSUSPENSAO").AsString
  vDestPes = DESTPES.FieldByName("ROTINACONTRATO").AsInteger
  If vPessoa >0 Then
    If vPesConcatena <>"" Then
      vPesConcatena = vPesConcatena + ","
    End If
    vPesConcatena = vPesConcatena + Str(vPessoa)
  End If
  DESTPES.Next
Wend
'Faz o teste para pessoa ou beneficiário e chama o relatório indicado no contrato
If vBenefConcatena <>"" Then
  'Se o relatório não estiver no contrato deve ser gerado um log desse beneficiário
  If vRelatorio <>"" Then
    QueryHandleRelatorio.Clear
    QueryHandleRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO =:PCODIGO")
    QueryHandleRelatorio.ParamByName("PCODIGO").AsString = vRelatorio
    QueryHandleRelatorio.Active = False
    QueryHandleRelatorio.Active = True

    ReportPreview(QueryHandleRelatorio.FieldByName("HANDLE").AsInteger, " A.ROTINACONTRATO = " + Str(vDestBen) + " AND A.BENEFICIARIO IN (" + vBenefConcatena + ")", True, False)
  Else
    BENE.Active = False
    BENE.Clear
    BENE.Add("SELECT NOME FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLEBENE")
    BENE.ParamByName("HANDLEBENE").Value = vBeneficiario
    BENE.Active = True
    vGravaOcorrencias = True
    vOcorrencias = vOcorrencias + Chr(13) + _
                   "Beneficiário: " + BENE.FieldByName("NOME").AsString + _
                   " | Contrato: " + DESTBEN.FieldByName("CONTRATO").AsString
  End If
End If
If vPesConcatena <>"" Then
  'Se o relatório não estiver no contrato deve ser gerado um log desse beneficiário
  If vRelatorio <>"" Then
    QueryHandleRelatorio.Clear
    QueryHandleRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO =:PCODIGO")
    QueryHandleRelatorio.ParamByName("PCODIGO").AsString = vRelatorio
    QueryHandleRelatorio.Active = False
    QueryHandleRelatorio.Active = True

    ReportPreview(QueryHandleRelatorio.FieldByName("HANDLE").AsInteger, " A.ROTINACONTRATO = " + Str(vDestPes) + " AND A.PESSOA IN (" + vPesConcatena + ")", True, False)
  Else
    PES.Active = False
    PES.Clear
    PES.Add("SELECT NOME FROM SFN_PESSOA WHERE HANDLE = :HANDLEPES")
    PES.ParamByName("HANDLEPES").Value = vPessoa
    PES.Active = True
    vGravaOcorrencias = True
    vOcorrencias = vOcorrencias + Chr(13) + _
                   "Pessoa: " + PES.FieldByName("NOME").AsString + _
                   " | Contrato: " + DESTPES.FieldByName("CONTRATO").AsString
  End If
End If

If vGravaOcorrencias Then
  'vOcorrencias =vOcorrencias +Chr(13)+"--------------------------------------------------------------------------------"
  Dim ATUALIZA As Object
  Set ATUALIZA = NewQuery

  If Not InTransaction Then StartTransaction

  ATUALIZA.Add("UPDATE SAM_ROTAVISOSUSPENSAO SET OCORRENCIAS = :OCORRENCIAS")
  ATUALIZA.Add(" WHERE HANDLE = :HANDLE")
  ATUALIZA.ParamByName("OCORRENCIAS").AsMemo = vOcorrencias
  ATUALIZA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  ATUALIZA.ExecSQL

  If InTransaction Then Commit

  Set ATUALIZA = Nothing
End If
Set PES = Nothing
Set BENE = Nothing
Set DESTPES = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim CONTAFIN As Object
  Dim dllGeraAvisoSusp As Object
  Dim DEST As Object
  Dim PROCESSA As Object
  Dim qRotinaSuspensao As Object
  Dim vgOLERestricaoFinan As Object
  Dim vContaFinanceira As Long
  Dim vUltimaConta As Long
  Dim vHandleDest As Long
  Dim vContrato As Long
  Dim vHandleRotinaContrato As Long
  Dim vTabResponsavel As Long
  Dim vHandleRotina As Long
  Dim vResponsavel As Long
  Dim vDias As Long
  Dim vDiasMin As Long
  Dim vDiasVencidos As Long
  Dim vDataAdesao As Date
  Dim vDataUltimoAviso As Date
  Dim vData As Date
  Dim vDataOutraRotina As Date
  Dim vrResponsavelTipo As Long
  Dim vrResponsavelHandle As Long
  Dim vrRestricaoDias As Long
  Dim vBenefResponsavel As Variant
  Dim vPessoaResponsavel As Variant
  Dim vCondicao As String
  Dim vrRestricaoTipo As String
  Dim CSProgress As Object
  Dim vEstado As Long
  Dim vMunicipio As Long
  Dim vBairro As String
  Dim vLogradouro As String
  Dim vNumero As Long
  Dim vComplemento As String
  Dim vCep As String
  Dim vIgualContrato As String
  Dim vBenef As Long
  Dim DiasVencidos As Long
  Dim vContinua As Boolean
  Dim vMensagemErro As String

  Set DEST = NewQuery
  Set PROCESSA = NewQuery
  Set qRotinaSuspensao = NewQuery
  Set CONTAFIN = CreateBennerObject("Financeiro.Contafin")
  Set dllGeraAvisoSusp = CreateBennerObject("SamAvisoSuspensao.GerarAviso")

  vHandleRotina = CurrentQuery.FieldByName("HANDLE").AsInteger
  vContinua = True

  'Verifica se já foi gerado alguma coisa
  DEST.Active = False
  DEST.Clear
  DEST.Add("SELECT DEST.HANDLE")
  DEST.Add("  FROM SAM_ROTAVISOSUSPENSAO_CONTRATO CONT,")
  DEST.Add("       SAM_ROTAVISOSUSPENSAO_DEST DEST")
  DEST.Add(" WHERE CONT.ROTINA = :ROTINA")
  DEST.Add("   AND CONT.HANDLE = DEST.ROTINACONTRATO")
  DEST.ParamByName("ROTINA").Value = vHandleRotina
  DEST.Active = True

  If Not InTransaction Then StartTransaction

  If Not DEST.EOF Then
    If MsgBox("Todos os dados já cadastrados serão apagados!" + Chr(13) + "Deseja Continuar?", 4) = vbYes Then
      Dim APAGAATRASO As Object
      Set APAGAATRASO = NewQuery


      APAGAATRASO.Active = False
      APAGAATRASO.Clear
      APAGAATRASO.Add("DELETE FROM SAM_ROTAVISOSUSPENSAO_ATRASO                   ")
      APAGAATRASO.Add(" WHERE EXISTS (Select A.HANDLE                             ")
      APAGAATRASO.Add("                    FROM SAM_ROTAVISOSUSPENSAO_ATRASO A,   ")
      APAGAATRASO.Add("                         SAM_ROTAVISOSUSPENSAO_DEST D,     ")
      APAGAATRASO.Add("                         SAM_ROTAVISOSUSPENSAO_CONTRATO C, ")
      APAGAATRASO.Add("                         SAM_ROTAVISOSUSPENSAO S           ")
      APAGAATRASO.Add("                   WHERE A.ROTINADEST = D.HANDLE           ")
      APAGAATRASO.Add("                     AND D.ROTINACONTRATO = C.HANDLE       ")
      APAGAATRASO.Add("                     And C.ROTINA = S.HANDLE               ")
      APAGAATRASO.Add("                     And A.HANDLE = SAM_ROTAVISOSUSPENSAO_ATRASO.HANDLE")
      APAGAATRASO.Add("                     And S.HANDLE = :HANDLEROTINA)         ")
      APAGAATRASO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      APAGAATRASO.ExecSQL

      Set DEST = Nothing

      Dim DESTINATARIO As Object
      Set DESTINATARIO = NewQuery

      DESTINATARIO.Active = False
      DESTINATARIO.Clear
      DESTINATARIO.Add("DELETE FROM SAM_ROTAVISOSUSPENSAO_DEST                  ")
      DESTINATARIO.Add(" WHERE EXISTS (Select D.HANDLE                          ")
      DESTINATARIO.Add("                 FROM SAM_ROTAVISOSUSPENSAO_DEST D,     ")
      DESTINATARIO.Add("                      SAM_ROTAVISOSUSPENSAO_CONTRATO C, ")
      DESTINATARIO.Add("                      SAM_ROTAVISOSUSPENSAO S           ")
      DESTINATARIO.Add("                WHERE D.ROTINACONTRATO = C.HANDLE       ")
      DESTINATARIO.Add("                  And C.ROTINA = S.HANDLE               ")
      DESTINATARIO.Add("                  And D.HANDLE = SAM_ROTAVISOSUSPENSAO_DEST.HANDLE")
      DESTINATARIO.Add("                  And S.HANDLE = :HANDLEROTINA)         ")
      DESTINATARIO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      DESTINATARIO.ExecSQL

      Set DEST = Nothing

    Else
      Exit Sub
    End If
  End If

  'Armazena todas as famílias dos contratos gravados
  qRotinaSuspensao.Active = False


    qRotinaSuspensao.Clear
    qRotinaSuspensao.Add("SELECT FAMILIA.CONTRATO,")
    qRotinaSuspensao.Add("       FAMILIA.DATAADESAO,")
    qRotinaSuspensao.Add("       FAMILIA.FAMILIA,")
    qRotinaSuspensao.Add("       FAMILIA.TABRESPONSAVEL,")
    qRotinaSuspensao.Add("       FAMILIA.TITULARRESPONSAVEL BENEFICIARIO,")
    qRotinaSuspensao.Add("       FAMILIA.PESSOARESPONSAVEL PESSOA,")
    qRotinaSuspensao.Add("       ROTINA.DATADOAVISO,")
    qRotinaSuspensao.Add("       ROTINA.DIAS,")
    qRotinaSuspensao.Add("       ROTINA.MESESDESCONSIDERAAVISO,")
    qRotinaSuspensao.Add("       C.DIASATRASO,")
    qRotinaSuspensao.Add("       CONTRATO.HANDLE HANDLEROTINACONTRATO,")
    qRotinaSuspensao.Add("       CONTABEN.HANDLE HANDLECONTAFINBEN,")
    qRotinaSuspensao.Add("       CONTAPES.HANDLE HANDLECONTAFINPES")
    qRotinaSuspensao.Add("  FROM SAM_ROTAVISOSUSPENSAO ROTINA,")
    qRotinaSuspensao.Add("       SAM_ROTAVISOSUSPENSAO_CONTRATO CONTRATO,")
    qRotinaSuspensao.Add("       SAM_CONTRATO C,")
    qRotinaSuspensao.Add("       SAM_FAMILIA FAMILIA")
    qRotinaSuspensao.Add(" LEFT JOIN SFN_CONTAFIN CONTABEN ON (FAMILIA.TITULARRESPONSAVEL = CONTABEN.BENEFICIARIO)")
    qRotinaSuspensao.Add(" LEFT JOIN SFN_CONTAFIN CONTAPES ON (FAMILIA.PESSOARESPONSAVEL = CONTAPES.PESSOA)")
    qRotinaSuspensao.Add(" WHERE ROTINA.HANDLE = CONTRATO.ROTINA")
    qRotinaSuspensao.Add("   AND FAMILIA.CONTRATO = CONTRATO.CONTRATO")
    qRotinaSuspensao.Add("       AND FAMILIA.CONTRATO = C.HANDLE")
    qRotinaSuspensao.Add("   AND ROTINA.HANDLE = :HANDLEROTINA")
    qRotinaSuspensao.Add("ORDER BY CONTABEN.HANDLE,")
    qRotinaSuspensao.Add("         CONTAPES.HANDLE,")
    qRotinaSuspensao.Add("         FAMILIA.CONTRATO,")
    qRotinaSuspensao.Add("         FAMILIA.FAMILIA")
    qRotinaSuspensao.ParamByName("HANDLEROTINA").Value = vHandleRotina

  qRotinaSuspensao.Active = True



  While Not qRotinaSuspensao.EOF 'para cada família é armazenado o seu responsável
    vTabResponsavel    = qRotinaSuspensao.FieldByName("TABRESPONSAVEL").AsInteger
    vBenefResponsavel  = 0
    vPessoaResponsavel = 0

    If (vTabResponsavel=1) Then
      vBenefResponsavel = qRotinaSuspensao.FieldByName("BENEFICIARIO").AsInteger
      vResponsavel      = vBenefResponsavel
      vBenef            = vBenefResponsavel

      'Busca a conta financeira para esse beneficiário
      vContaFinanceira = CONTAFIN.Qual(CurrentSystem, vBenefResponsavel, 1)

      'Procura em outras rotinas esse beneficiário
      Dim PROCURA As Object
      Set PROCURA = NewQuery
      PROCURA.Add("Select S.DATADOAVISO                     ")
      PROCURA.Add("  FROM SAM_ROTAVISOSUSPENSAO_DEST D,     ")
      PROCURA.Add("       SAM_ROTAVISOSUSPENSAO_CONTRATO C, ")
      PROCURA.Add("       SAM_ROTAVISOSUSPENSAO S           ")
      PROCURA.Add(" WHERE D.ROTINACONTRATO = C.HANDLE       ")
      PROCURA.Add("   And C.ROTINA = S.HANDLE               ")
      PROCURA.Add("   And S.HANDLE <> :HANDLEROTINA         ")
      PROCURA.Add("   And D.BENEFICIARIO = :HANDLEBENEF     ")
      PROCURA.Add(" ORDER BY S.DATADOAVISO DESC             ")
      PROCURA.ParamByName("HANDLEROTINA").Value = vHandleRotina
      PROCURA.ParamByName("HANDLEBENEF").Value = vBenefResponsavel
      PROCURA.Active = True

      vDataOutraRotina = PROCURA.FieldByName("DATADOAVISO").AsDateTime
      Set PROCURA = Nothing

      Dim ENDBEN As Object
      Set ENDBEN = NewQuery
      ENDBEN.Add("Select E.ESTADO,")
      ENDBEN.Add("       E.MUNICIPIO,")
      ENDBEN.Add("       E.BAIRRO,")
      ENDBEN.Add("       E.LOGRADOURO,")
      ENDBEN.Add("       E.NUMERO,")
      ENDBEN.Add("       E.COMPLEMENTO,")
      ENDBEN.Add("       E.CEP")
      ENDBEN.Add("  FROM SAM_BENEFICIARIO B,")
      ENDBEN.Add("       SAM_ENDERECO E")
      ENDBEN.Add(" WHERE B.ENDERECOCORRESPONDENCIA = E.HANDLE")
      ENDBEN.Add("   And B.HANDLE = :HANDLEBENEF")
      ENDBEN.Add("   AND (B.DATACANCELAMENTO IS NULL OR B.DATACANCELAMENTO > :DATABASE)")
      ENDBEN.ParamByName("DATABASE").AsDateTime = CurrentQuery.FieldByName("DATADOAVISO").AsDateTime
      ENDBEN.ParamByName("HANDLEBENEF").Value = vBenefResponsavel
      ENDBEN.Active = True

      vContinua = (Not ENDBEN.EOF)
      vEstado = ENDBEN.FieldByName("ESTADO").AsInteger
      vMunicipio = ENDBEN.FieldByName("MUNICIPIO").AsInteger
      vBairro = ENDBEN.FieldByName("BAIRRO").AsString
      vLogradouro = ENDBEN.FieldByName("LOGRADOURO").AsString
      vNumero = ENDBEN.FieldByName("NUMERO").AsInteger
      vComplemento = ENDBEN.FieldByName("COMPLEMENTO").AsString
      vCep = ENDBEN.FieldByName("CEP").AsString
      Set ENDBEN = Nothing

    ElseIf (vTabResponsavel=2) Then
      'DANIELA -passar sempre o beneficiário pra descobrir os dias de atraso...
      Dim q1 As Object
      Set q1 = NewQuery
      Dim q2 As Object
      Set q2 = NewQuery
      vPessoaResponsavel = qRotinaSuspensao.FieldByName("PESSOA").AsInteger
      vResponsavel       = vPessoaResponsavel
      q1.Clear
      q1.Add("Select * from sam_contrato where pessoa = :PESSOA")
      q1.ParamByName("PESSOA").AsInteger = vPessoaResponsavel
      q1.Active = True

      If Not q1.EOF Then
        q2.Clear
        q2.Add("SELECT B.HANDLE, C.DIASRESTRICAOFINANCEIRA, C.DATAADESAO")
        q2.Add("       FROM SAM_CONTRATO C, SAM_BENEFICIARIO B")
        q2.Add("WHERE  C.HANDLE = B.CONTRATO AND C.PESSOA = :PESSOA")
        q2.Add("       AND B.DATACANCELAMENTO IS NULL")
        q2.Add("       AND C.DATACANCELAMENTO IS NULL")
        q2.ParamByName("PESSOA").AsInteger = qRotinaSuspensao.FieldByName("PESSOA").AsInteger
        q2.Active = True
        vBenef = q2.FieldByName("HANDLE").AsInteger
      Else
        q1.Clear
        q1.Add("Select * from sam_FAMILIA where pessoaRESPONSAVEL = :PESSOA")
        q1.ParamByName("PESSOA").AsInteger = vPessoaResponsavel
        q1.Active = True

        q2.Clear
        q2.Add("Select B.HANDLE, C.DIASRESTRICAOFINANCEIRA, C.DATAADESAO")
        q2.Add("  FROM SAM_CONTRATO C, SAM_FAMILIA F, SAM_BENEFICIARIO B")
        q2.Add(" WHERE F.CONTRATO = C.HANDLE")
        q2.Add("   And F.HANDLE = B.FAMILIA")
        q2.Add("   And F.PESSOARESPONSAVEL = :PESSOA")
        q2.Add("   And B.DATACANCELAMENTO Is Null")
        q2.Add("   And C.DATACANCELAMENTO Is Null")
        q2.ParamByName("PESSOA").AsInteger = vPessoaResponsavel
        q2.Active = True
        vBenef = q2.FieldByName("HANDLE").AsInteger
      End If

      'Busca a conta financeira para esse beneficiário
      vContaFinanceira = CONTAFIN.Qual(CurrentSystem, vPessoaResponsavel, 3)

      'Procura em outras rotinas essa Pessoa
      Dim PROCURADATA As Object
      Set PROCURADATA = NewQuery
      PROCURADATA.Add("Select S.DATADOAVISO                     ")
      PROCURADATA.Add("  FROM SAM_ROTAVISOSUSPENSAO_DEST D,     ")
      PROCURADATA.Add("       SAM_ROTAVISOSUSPENSAO_CONTRATO C, ")
      PROCURADATA.Add("       SAM_ROTAVISOSUSPENSAO S           ")
      PROCURADATA.Add(" WHERE D.ROTINACONTRATO = C.HANDLE       ")
      PROCURADATA.Add("   And C.ROTINA = S.HANDLE               ")
      PROCURADATA.Add("   And S.HANDLE <> :HANDLEROTINA         ")
      PROCURADATA.Add("   And D.PESSOA = :HANDLEPESSOA          ")
      PROCURADATA.Add(" ORDER BY S.DATADOAVISO DESC             ")
      PROCURADATA.ParamByName("HANDLEROTINA").Value = vHandleRotina
      PROCURADATA.ParamByName("HANDLEPESSOA").Value = vPessoaResponsavel
      PROCURADATA.Active = True

      vDataOutraRotina = PROCURADATA.FieldByName("DATADOAVISO").AsDateTime
      Set PROCURADATA = Nothing

      Dim ENDPES As Object
      Set ENDPES = NewQuery
      ENDPES.Add("Select E.ESTADO,")
      ENDPES.Add("       E.MUNICIPIO,")
      ENDPES.Add("       E.BAIRRO,")
      ENDPES.Add("		 E.LOGRADOURO,")
      ENDPES.Add("		 E.NUMERO,")
      ENDPES.Add("		 E.COMPLEMENTO,")
      ENDPES.Add("       E.CEP")
      ENDPES.Add("  FROM SFN_PESSOA E")
      ENDPES.Add(" WHERE E.HANDLE = :HANDLEPESSOA")
      ENDPES.ParamByName("HANDLEPESSOA").Value = vPessoaResponsavel
      ENDPES.Active = True

      vEstado = ENDPES.FieldByName("ESTADO").AsInteger
      vMunicipio = ENDPES.FieldByName("MUNICIPIO").AsInteger
      vBairro = ENDPES.FieldByName("BAIRRO").AsString
      vLogradouro = ENDPES.FieldByName("LOGRADOURO").AsString
      vNumero = ENDPES.FieldByName("NUMERO").AsInteger
      vComplemento = ENDPES.FieldByName("COMPLEMENTO").AsString
      vCep = ENDPES.FieldByName("CEP").AsString
      Set ENDPES = Nothing
    End If

	If vContinua Then
	    vDias = qRotinaSuspensao.FieldByName("DIAS").AsInteger
	    vDiasMin = qRotinaSuspensao.FieldByName("DIASATRASO").AsInteger
	    vDataAdesao = qRotinaSuspensao.FieldByName("DATAADESAO").AsDateTime
	    vData = qRotinaSuspensao.FieldByName("DATADOAVISO").AsDateTime
	    vDataUltimoAviso = DateAdd("m", qRotinaSuspensao.FieldByName("MESESDESCONSIDERAAVISO").AsInteger * -1, qRotinaSuspensao.FieldByName("DATADOAVISO").AsDateTime)
	    vHandleRotinaContrato = qRotinaSuspensao.FieldByName("HANDLEROTINACONTRATO").AsInteger
	    vContrato = qRotinaSuspensao.FieldByName("CONTRATO").AsInteger


	    'Depois de armazenado os dados da rotina é chamada uma função para fazer o cálcula da qtde de dias em atraso
	    CONTAFIN.Restricao(CurrentSystem, vBenef, vDias, vDataAdesao, vData, vrResponsavelTipo, vrResponsavelHandle, vrRestricaoTipo, vrRestricaoDias)

	    'Verifica se os dias em atraso esta no intervalo mínimo do contrato e o informado na rotina
	    'Verifica também se existe outra rotina com esse destinatário cadastrado antes da data informada na rotina
	    If CurrentQuery.FieldByName("TABDIAS").AsInteger = 1 Then
	      vDias = vDiasMin
	    End If

	    If(vrRestricaoDias >vDias)And _
	      (vDataOutraRotina <vDataUltimoAviso)And _
	      (vUltimaConta <>vContaFinanceira)Then

	      vHandleDest = NewHandle("SAM_ROTAVISOSUSPENSAO_DEST")
	      vUltimaConta = vContaFinanceira

		  If dllGeraAvisoSusp.FinalizaAvisoSuspensao(CurrentSystem, vTabResponsavel, vHandleDest, vContrato, vHandleRotinaContrato, _
	                                      vResponsavel, vEstado, vMunicipio, vBairro,vLogradouro, vComplemento, vNumero, _
	                                      vCep, vContaFinanceira, vData, vDias, vMensagemErro) = 1 Then

		    bsShowMessage("Erro ao tentar finalizar a rotina de aviso de suspensão: " + vMensagemErro, "E")
		    GoTo lbFinalizar
		    Exit Sub
		  End If

	    End If
    End If

  qRotinaSuspensao.Next
Wend

PROCESSA.Active = False
PROCESSA.Clear
PROCESSA.Add("UPDATE SAM_ROTAVISOSUSPENSAO")
PROCESSA.Add("   Set PROCESSADO = 'S'")
PROCESSA.Add(" WHERE HANDLE = :HANDLEROTINA")
PROCESSA.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
PROCESSA.ExecSQL

If InTransaction Then Commit

'MsgBox("Processo Concluído !")

WriteAudit("P", HandleOfTable("SAM_ROTAVISOSUSPENSAO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Aviso de Suspensão - Processamento")

RefreshNodesWithTable("SAM_ROTAVISOSUSPENSAO")

GoTo lbFinalizar

lbFinalizar:
  Set qRotinaSuspensao = Nothing
  Set PROCESSA = Nothing
  Set q1 = Nothing
  Set dllGeraAvisoSusp = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  'Dim PARAMETROS As Object
  'Set PARAMETROS =NewQuery
  '  PARAMETROS.Active =False
  '  PARAMETROS.Add("SELECT DIASPARAAVISO")
  '  PARAMETROS.Add("  FROM SAM_PARAMETROSBENEFICIARIO")
  '  PARAMETROS.Active =True


  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser
  CurrentQuery.FieldByName("DATADOAVISO").Value = ServerDate
  'CurrentQuery.FieldByName("DIAS").Value=PARAMETROS.FieldByName("DIASPARAAVISO").AsInteger

  'Set PARAMETROS =Nothing
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("PROCESSADO").AsString = "S" Then
    BOTAOPROCESSAR.Caption = "Reprocessar"
  Else
    BOTAOPROCESSAR.Caption = "Processar"
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim DEST As Object
  Dim CONTRATO As Object
  Set DEST = NewQuery
  Set CONTRATO = NewQuery

  Dim APAGAATRASO As Object
  Set APAGAATRASO = NewQuery

  APAGAATRASO.Active = False
  APAGAATRASO.Clear
  APAGAATRASO.Add("DELETE FROM SAM_ROTAVISOSUSPENSAO_ATRASO                ")
  APAGAATRASO.Add(" WHERE EXISTS (Select A.HANDLE                          ")
  APAGAATRASO.Add("                 FROM SAM_ROTAVISOSUSPENSAO_ATRASO A,   ")
  APAGAATRASO.Add("                      SAM_ROTAVISOSUSPENSAO_DEST D,     ")
  APAGAATRASO.Add("                      SAM_ROTAVISOSUSPENSAO_CONTRATO C, ")
  APAGAATRASO.Add("                      SAM_ROTAVISOSUSPENSAO S           ")
  APAGAATRASO.Add("                WHERE A.ROTINADEST = D.HANDLE           ")
  APAGAATRASO.Add("                  AND D.ROTINACONTRATO = C.HANDLE       ")
  APAGAATRASO.Add("                  And C.ROTINA = S.HANDLE               ")
  APAGAATRASO.Add("                  And A.HANDLE = SAM_ROTAVISOSUSPENSAO_ATRASO.HANDLE")
  APAGAATRASO.Add("                  And S.HANDLE = :HANDLEROTINA)         ")
  APAGAATRASO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  APAGAATRASO.ExecSQL

  DEST.Active = False
  DEST.Clear
  DEST.Add("DELETE FROM SAM_ROTAVISOSUSPENSAO_DEST                   ")
  DEST.Add(" WHERE EXISTS (Select D.HANDLE                           ")
  DEST.Add("                 FROM SAM_ROTAVISOSUSPENSAO_DEST D,      ")
  DEST.Add("                      SAM_ROTAVISOSUSPENSAO_CONTRATO C,  ")
  DEST.Add("       		          SAM_ROTAVISOSUSPENSAO S            ")
  DEST.Add("                WHERE D.ROTINACONTRATO = C.HANDLE        ")
  DEST.Add("                  And C.ROTINA = S.HANDLE                ")
  DEST.Add("                  And D.HANDLE = SAM_ROTAVISOSUSPENSAO_DEST.HANDLE")
  DEST.Add("                  And S.HANDLE = :HANDLEROTINA)          ")
  DEST.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DEST.ExecSQL

  CONTRATO.Active = False
  CONTRATO.Clear
  CONTRATO.Add("DELETE FROM SAM_ROTAVISOSUSPENSAO_CONTRATO                  ")
  CONTRATO.Add(" WHERE EXISTS (SELECT ROTCONT.HANDLE                        ")
  CONTRATO.Add("                 FROM SAM_ROTAVISOSUSPENSAO          ROT,   ")
  CONTRATO.Add("                      SAM_ROTAVISOSUSPENSAO_CONTRATO ROTCONT")
  CONTRATO.Add("                WHERE ROT.HANDLE = ROTCONT.ROTINA           ")
  CONTRATO.Add("                  And ROTCONT.HANDLE = SAM_ROTAVISOSUSPENSAO_CONTRATO.HANDLE")
  CONTRATO.Add("                  And ROT.HANDLE = :HANDLEROTINA)           ")
  CONTRATO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  CONTRATO.ExecSQL

  Set APAGAATRASO = Nothing
  Set DEST = Nothing
  Set CONTRATO = Nothing
End Sub

