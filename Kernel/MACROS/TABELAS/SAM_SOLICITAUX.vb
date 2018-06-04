'HASH: FBF3FF73DFF892BD0BA609EC8D75CD93
'SAM_SOLICITAUX
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BENEFICIARIO_OnExit()
  Dim SQLBEN As Object
  Dim SQLFAM As Object
  Dim SQLTIT As Object
  Dim SQLCON As Object
  Dim SQLPES As Object
  Set SQLPES = NewQuery
  Set SQLCON = NewQuery
  Set SQLBEN = NewQuery
  Set SQLFAM = NewQuery
  Set SQLTIT = NewQuery

  SQLBEN.Active = False
  SQLBEN.Clear
  SQLBEN.Add("SELECT NOME,FAMILIA,CONTRATO FROM SAM_BENEFICIARIO")
  SQLBEN.Add("WHERE HANDLE = :HBENEFICIARIO")
  SQLBEN.ParamByName("HBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQLBEN.Active = True

  SQLCON.Active = False
  SQLCON.Clear
  SQLCON.Add("SELECT CONTRATO,CNPJ,CPFRESPONSAVEL,PESSOA FROM SAM_CONTRATO WHERE HANDLE = :HCONTRATO")
  SQLCON.ParamByName("HCONTRATO").AsInteger = SQLBEN.FieldByName("CONTRATO").AsInteger
  SQLCON.Active = True

  SQLFAM.Active = False
  SQLFAM.Clear
  SQLFAM.Add("SELECT TITULARRESPONSAVEL,TABRESPONSAVEL FROM SAM_FAMILIA WHERE HANDLE = :HFAMILIA")
  SQLFAM.ParamByName("HFAMILIA").AsInteger = SQLBEN.FieldByName("FAMILIA").AsInteger
  SQLFAM.Active = True

  SQLTIT.Active = False
  SQLTIT.Clear
  SQLTIT.Add("SELECT NOME FROM SAM_BENEFICIARIO WHERE HANDLE = :HBENEFICIARIO")
  SQLTIT.ParamByName("HBENEFICIARIO").AsInteger = SQLFAM.FieldByName("TITULARRESPONSAVEL").AsInteger
  SQLTIT.Active = True

  SQLPES.Active = False
  SQLPES.Clear
  SQLPES.Add("SELECT CNPJCPF, NOME FROM SFN_PESSOA WHERE HANDLE = :HPESSOA")
  SQLPES.ParamByName("HPESSOA").AsInteger = SQLCON.FieldByName("PESSOA").AsInteger
  SQLPES.Active = True

  If SQLBEN.EOF Then
    ROTULOBEN01.Text = "*** BENEFICIARIO NÃO ENCONTRADO ***"
    ROTULOBEN02.Text = "*"
  Else
    If SQLFAM.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then ' se o responsável for beneficiário
      ROTULOBEN01.Text = "Beneficiário : " + SQLBEN.FieldByName("NOME").AsString
      ROTULOBEN02.Text = "Titilar Responsável - Beneficiario: " + Format(SQLCON.FieldByName("CPFRESPONSAVEL").AsString, "000\.000\.000\-00") + " - " + SQLTIT.FieldByName("NOME").AsString
    Else 'Caso o Beneficiario for Pessoa
      ROTULOBEN01.Text = "Beneficiario : " + SQLBEN.FieldByName("NOME").AsString 'Caso o responsável for pessoa
      ROTULOBEN02.Text = "Titular Responsável - Pessoa: " + Format(SQLCON.FieldByName("CNPJ").AsString, "00\.000\.000\/0000\-00") + " - " + SQLTIT.FieldByName("NOME").AsString
    End If
  End If

  If CurrentQuery.State <>1 And _
      Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull Then
    Dim SQL As Object
    Dim vEndereco As Long
    Dim vEstado As Long
    Dim vMunicipio As Long

    Set SQL = NewQuery

    SQL.Clear



      SQL.Add("SELECT B.ENDERECORESIDENCIAL HBENEF_ENDERECO,")
      SQL.Add("       C.ESTADO CONT_ESTADO, C.MUNICIPIO CONT_MUNICIPIO,")
      SQL.Add("       FAM.TABRESPONSAVEL,")
      SQL.Add("       TIT.ENDERECORESIDENCIAL HTIT_ENDERECO,")
      SQL.Add("       PES.ESTADO PES_ESTADO, PES.MUNICIPIO PES_MUNICIPIO")
      SQL.Add("FROM SAM_BENEFICIARIO B")
      SQL.Add("     JOIN SAM_CONTRATO C ON")
      SQL.Add("          C.HANDLE = B.CONTRATO")
      SQL.Add("     JOIN SAM_FAMILIA FAM ON")
      SQL.Add("     (FAM.HANDLE = B.FAMILIA)")
      SQL.Add("     LEFT JOIN SAM_BENEFICIARIO TIT ON")
      SQL.Add("     (TIT.HANDLE = FAM.TITULARRESPONSAVEL)")
      SQL.Add("     LEFT JOIN SFN_PESSOA PES ON")
      SQL.Add("     (PES.HANDLE = FAM.PESSOARESPONSAVEL)")
      SQL.Add("WHERE B.HANDLE = :HBENEFICIARIO")


    SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    SQL.Active = True

    vEndereco = SQL.FieldByName("HBENEF_ENDERECO").AsInteger
    If vEndereco = 0 Then
      If SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
        vEndereco = SQL.FieldByName("HTIT_ENDERECO").AsInteger
      Else
        vEstado = SQL.FieldByName("PES_ESTADO").AsInteger
        vMunicipio = SQL.FieldByName("PES_MUNICIPIO").AsInteger
      End If
    End If

    If vEndereco = 0 And _
                   vEstado = 0 Then
      vEstado = SQL.FieldByName("CONT_ESTADO").AsInteger
      vMunicipio = SQL.FieldByName("CONT_MUNICIPIO").AsInteger
    End If

    If vEstado = 0 Then
      Dim SQLEndereco As Object
      Set SQLEndereco = NewQuery

      SQLEndereco.Clear
      SQLEndereco.Add("SELECT ESTADO, MUNICIPIO")
      SQLEndereco.Add("FROM SAM_ENDERECO")
      SQLEndereco.Add("WHERE HANDLE = :HENDERECO")
      SQLEndereco.ParamByName("HENDERECO").Value = vEndereco
      SQLEndereco.Active = True

      vEstado = SQLEndereco.FieldByName("ESTADO").AsInteger
      vMunicipio = SQLEndereco.FieldByName("MUNICIPIO").AsInteger

      Set SQLEndereco = Nothing
    End If

    If vEstado <>0 Then
      CurrentQuery.FieldByName("ESTADO").Value = vEstado
      If vMunicipio <>0 Then
        CurrentQuery.FieldByName("MUNICIPIO").Value = vMunicipio
      End If
    End If

    Set SQL = Nothing
  End If

  SQLBEN.Active = False
  SQLFAM.Active = False
  SQLTIT.Active = False
  SQLCON.Active = False
  SQLPES.Active = False
  Set SQLPES = Nothing
  Set SQLCON = Nothing
  Set SQLBEN = Nothing
  Set SQLFAM = Nothing
  Set SQLTIT = Nothing
End Sub

'#Uses "*ProcuraBeneficiarioAtivo"

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraBeneficiarioAtivo(True, ServerDate, BENEFICIARIO.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = vHandle
  End If
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  'SMS 82885 - Artur - Bloqueio de botões quando o registro está em modo de edição
  Dim SQLUPD As Object
  If CurrentQuery.State > 1 Then
    BsShowMessage("O registro está em edição", "I")
  	Exit Sub
  End If

  Dim SQLUSER As Object
  Set SQLUSER = NewQuery

  SQLUSER.Clear
  SQLUSER.Add("SELECT C.GUIA, D.PEG ")
  SQLUSER.Add(" FROM SAM_SOLICITAUX_BENEFICIO_GUIA A")
  SQLUSER.Add(" JOIN SAM_SOLICITAUX_BENEFICIO E ON (A.SOLICITAUXBENEFICIO = E.HANDLE)")
  SQLUSER.Add(" JOIN SAM_GUIA_EVENTOS B ON (A.GUIAEVENTO = B.HANDLE) ")
  SQLUSER.Add(" JOIN SAM_GUIA C ON (B.GUIA = C.HANDLE)")
  SQLUSER.Add(" JOIN SAM_PEG  D On (C.PEG = D.HANDLE)")
  SQLUSER.Add(" WHERE E.SOLICITAUX = :HANDLE")
  SQLUSER.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  SQLUSER.Active = True

  If Not SQLUSER.EOF Then
 	BsShowMessage("Não é possível cancelar, pois a solicitação está vinculada à guia: " + SQLUSER.FieldByName("GUIA").AsString + " PEG: " +  SQLUSER.FieldByName("PEG").AsString, "I")
    Set SQLUSER = Nothing
    Exit Sub
  End If

  SQLUSER.Active = False
  SQLUSER.Clear
  SQLUSER.Add("SELECT NOME FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HUSUARIO")
  SQLUSER.ParamByName("HUSUARIO").AsInteger = CurrentUser
  SQLUSER.Active = True

  If CurrentQuery.FieldByName("SITUACAO").AsString = "C" Then
    BsShowMessage("Solicitação já cancelada", "I")
    Set SQLUSER = Nothing
    Exit Sub
  End If

  'Só cancela se ela não estiver negada
  If CurrentQuery.FieldByName("SITUACAO").AsString <>"N" Then
    If BsShowMessage("Deseja cancelar?", "Q") = vbYes Then

      Set SQLUPD = NewQuery

      If Not InTransaction Then StartTransaction

      SQLUPD.Clear
      SQLUPD.Add("UPDATE SAM_SOLICITAUX SET")
      SQLUPD.Add("  OCORRENCIAS = :OCORRENCIAS,")
      SQLUPD.Add("  SITUACAO = 'C'")
      SQLUPD.Add("WHERE HANDLE = :HSOLICITAUX")
      SQLUPD.ParamByName("HSOLICITAUX").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Solicitação Cancelada"
      SQLUPD.ExecSQL

      If InTransaction Then Commit

      Set SQLUPD = Nothing

      SQLUSER.Active = False
      Set SQLUSER = Nothing
      RefreshNodesWithTable("SAM_SOLICITAUX")
    End If
  Else
  	BsShowMessage("Não é possível cancelar uma solicitação negada", "I")
  End If
End Sub

Public Sub BOTAOLIBERAR_OnClick()
  'SMS 82885 - Artur - Bloqueio de botões quando o registro está em modo de edição
  Dim SQLUPD As Object
  If CurrentQuery.State > 1 Then
	BsShowMessage("O registro está em edição", "I")
  	Exit Sub
  End If


  Dim SQL As Object
  Dim SQLDOC As Object
  Dim SQLPREST As Object
  Dim SQLPAC As Object
  Dim SQLBEN As Object
  Dim SQLUSER As Object
  Dim vLimiteAdiantamentoDefinitivo As Double

  Set SQL = NewQuery
  Set SQLDOC = NewQuery
  Set SQLPREST = NewQuery
  Set SQLPAC = NewQuery
  Set SQLBEN = NewQuery
  Set SQLUSER = NewQuery

  If Not InTransaction Then StartTransaction

  'André - SMS 28326 - 13/10/2004
  SQLDOC.Clear
  SQLDOC.Add (" SELECT B.HANDLE                               ")
  SQLDOC.Add ("   FROM SAM_SOLICITAUX_BENEFICIO_DOC D,        ")
  SQLDOC.Add ("        SAM_SOLICITAUX_BENEFICIO B             ")
  SQLDOC.Add ("  WHERE B.SOLICITAUX = :SOLICIT                ")
  SQLDOC.Add ("    AND B.HANDLE = D.SOLICITAUXBENEFICIO       ")
  SQLDOC.ParamByName("SOLICIT").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQLDOC.Active = True

  If SQLDOC.FieldByName("HANDLE").IsNull Then
	BsShowMessage("Não é possível liberar a solicitação de auxílio sem documento", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABCLASSIFICACAO").Value = 2 Then

    SQLPREST.Clear
    SQLPREST.Add (" SELECT B.HANDLE, B.FATURAPESSOAAUXILIO       ")
    SQLPREST.Add ("   FROM SAM_SOLICITAUX A,                     ")
    SQLPREST.Add ("        SAM_SOLICITAUX_BENEFICIO B            ")
    SQLPREST.Add ("  WHERE A.HANDLE = B.SOLICITAUX               ")
    SQLPREST.Add ("    AND A.TABCLASSIFICACAO = 2                ")
    SQLPREST.Add ("    AND A.SITUACAO IN ('T','L')               ")
    SQLPREST.Add ("    AND B.FATURAPESSOAAUXILIO IS NULL         ")
    SQLPREST.Add ("    AND A.BENEFICIARIO = :BENEF               ")
    SQLPREST.Add ("    AND A.HANDLE NOT IN (:SOLICIT)            ")
    SQLPREST.ParamByName("BENEF").Value = CurrentQuery.FieldByName("BENEFICIARIO").Value
    SQLPREST.ParamByName("SOLICIT").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQLPREST.Active = True

    If Not (SQLPREST.FieldByName("HANDLE").IsNull) Then
	  BsShowMessage("Não é possível liberar a solicitação, pois existe fatura que não foi prestada conta ainda.", "I")
      Exit Sub
    End If

  End If

  'FIM SMS 28326


  If CurrentQuery.FieldByName("SITUACAO").AsString = "L" Then 'Se já estiver liberada
    BsShowMessage("Solicitação já liberada", "I")
    Exit Sub
  Else

    If CurrentQuery.FieldByName("TABCLASSIFICACAO").AsInteger = 1 Then
      SQL.Clear
      SQL.Add("SELECT HANDLE")
      SQL.Add("FROM SAM_SOLICITAUX_BENEFICIO")
      SQL.Add("WHERE SOLICITAUX = :HSOLICITAUX")
      SQL.Add("  AND (   QUANTIDADE IS NULL")
      SQL.Add("       OR VALOREVENTO IS NULL")
      SQL.Add("       OR VALORUNITARIOLIBERADO IS NULL)")
      SQL.ParamByName("HSOLICITAUX").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL.Active = True

      If Not SQL.EOF Then
        BsShowMessage("Os valores dos benefícios não foram informados", "I")
        Exit Sub
      End If
    End If

    SQLUSER.Active = False
    SQLUSER.Clear
    SQLUSER.Add("SELECT NOME FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HUSUARIO")
    SQLUSER.ParamByName("HUSUARIO").AsInteger = CurrentUser
    SQLUSER.Active = True

    'Busca o nº do contrato na tabela de Beneficiarios
    SQLBEN.Active = False
    SQLBEN.Clear
    SQLBEN.Add(" SELECT CONTRATO,DATAADMISSAO FROM SAM_BENEFICIARIO")
    SQLBEN.Add(" WHERE HANDLE = :BENEFICIARIO")
    SQLBEN.ParamByName("BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    SQLBEN.Active = True

    'Verifica se a data de cadastro está dentro da vigência do contrato
    SQLPAC.Active = False
    SQLPAC.Clear
    SQLPAC.Add(" SELECT DATAINICIAL, DATAFINAL, LIMITEADIANTAMENTODEF FROM SAM_CONTRATO_AUXILIO ")
    SQLPAC.Add("  WHERE CONTRATO = :CONTRATOBEN                         ")
    'Balani SMS 54421 17/01/2005
    SQLPAC.Add("    AND DATAINICIAL <= :DATAADMISSAO")
    SQLPAC.Add("    AND (DATAFINAL IS NULL OR DATAFINAL >= :DATAADMISSAO)")
    SQLPAC.ParamByName("DATAADMISSAO").AsDateTime = SQLBEN.FieldByName("DATAADMISSAO").AsDateTime
    'final SMS 54421
    SQLPAC.ParamByName("CONTRATOBEN").AsInteger = SQLBEN.FieldByName("CONTRATO").AsInteger
    SQLPAC.Active = True

    Set SQLUPD = NewQuery

    SQLUPD.Clear
    SQLUPD.Add("UPDATE SAM_SOLICITAUX SET")
    SQLUPD.Add("  OCORRENCIAS = :OCORRENCIAS")
    SQLUPD.Add("WHERE HANDLE = :HSOLICITAUX")
    SQLUPD.ParamByName("HSOLICITAUX").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

    'Caso o select não retornar nada
    If SQLPAC.EOF Then
      SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio negado. O contrato não posui configuração de auxílio vigente para esta data de solicitação"
      SQLUPD.ExecSQL

      Set SQLUPD = Nothing

	  BsShowMessage("Auxílio negado. O contrato não posui configuração de auxílio", "I")
      RefreshNodesWithTable("SAM_SOLICITAUX")
      Exit Sub
    Else
      'Verifica se o Beneficiário tem a DATAADMISSAO entre o período válido
      If((SQLBEN.FieldByName("DATAADMISSAO").AsDateTime <SQLPAC.FieldByName("DATAINICIAL").AsDateTime)And(Not SQLPAC.FieldByName("DATAINICIAL").IsNull))Or _
         ((SQLBEN.FieldByName("DATAADMISSAO").AsDateTime >SQLPAC.FieldByName("DATAFINAL").AsDateTime)And(Not SQLPAC.FieldByName("DATAFINAL").IsNull))Then
      SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio negado. Data de admissão fora do período de vigência"

      SQLUPD.ExecSQL

      Set SQLUPD = Nothing

	  BsShowMessage("Auxílio Negado. Data de Admissão fora do período permitido nos parâmetros de auxílio do contrato", "I")
      RefreshNodesWithTable("SAM_SOLICITAUX")
      Exit Sub
    End If
  End If

  vLimiteAdiantamentoDefinitivo = SQLPAC.FieldByName("LIMITEADIANTAMENTODEF").AsFloat

  'Verifica se o campo NAOPERMITEAUXILIO está marcado
  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT NAOPERMITEAUXILIO FROM SAM_BENEFICIARIO WHERE HANDLE = :HBeneficiario ")
  SQL.ParamByName("HBeneficiario").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("NAOPERMITEAUXILIO").AsString = "S" Then
    SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio negado. Beneficiário não permite auxílio "

    SQLUPD.ExecSQL

    Set SQLUPD = Nothing

	BsShowMessage("Auxílio Negado. Beneficiário não permite auxílio", "I")
    SQL.Active = False
    Set SQL = Nothing
    RefreshNodesWithTable("SAM_SOLICITAUX")
    Exit Sub
  End If

  'Verifica se existe o Beneficiário está de Licença
  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT BEN.DATAADMISSAO                           ")
  SQL.Add("  FROM SAM_BENEFICIARIO BEN,                      ")
  SQL.Add("       SAM_BENEFICIARIO_LICENCA AUX               ")
  SQL.Add(" WHERE AUX.BENEFICIARIO = BEN.HANDLE              ")
  SQL.Add("   AND BEN.HANDLE =:HBENEFICIARIO                 ")
  SQL.Add("   AND ((BEN.DATAADMISSAO >= AUX.DATAINICIAL      ")
  SQL.Add("   AND BEN.DATAADMISSAO <= AUX.DATAFINAL)         ")
  SQL.Add("    OR (AUX.DATAFINAL IS NULL                     ")
  SQL.Add("   AND AUX.DATAINICIAL IS NOT NULL)               ")
  SQL.Add("   AND (AUX.PARTICULAR = 'S'))                    ")
  SQL.ParamByName("HBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio negado. Beneficiario está de Licença particular"

    SQLUPD.ExecSQL

    Set SQLUPD = Nothing

	BsShowMessage("Auxílio Negado. Beneficiário de Licença particular", "I")
    SQL.Active = False
    Set SQL = Nothing
    RefreshNodesWithTable("SAM_SOLICITAUX")
    Exit Sub
  End If

  'Verificar a classificação do pacote
  If Not CurrentQuery.FieldByName("PACOTEAUXILIO").IsNull Then
    Dim vPacoteLimiteTipo As String
    Dim vPacoteLimiteValor As Double
    Dim vTotalBeneficios As Double
    Dim vFilial As Long

    SQL.Clear
    SQL.Add("SELECT LIMITETIPO, LIMITEVALOR")
    SQL.Add("FROM SAM_PACOTEAUXILIO")
    SQL.Add("WHERE HANDLE = :HPACOTEAUXILIO")
    SQL.ParamByName("HPACOTEAUXILIO").Value = CurrentQuery.FieldByName("PACOTEAUXILIO").AsInteger
    SQL.Active = True

    vPacoteLimiteTipo = SQL.FieldByName("LIMITETIPO").AsString
    vPacoteLimiteValor = SQL.FieldByName("LIMITEVALOR").AsFloat

    SQL.Clear
    SQL.Add("SELECT SUM(QUANTIDADE * VALORUNITARIOLIBERADO) TOTALBENEFICIOS")
    SQL.Add("FROM SAM_SOLICITAUX_BENEFICIO")
    SQL.Add("WHERE SOLICITAUX = :HSOLICITAUX")
    SQL.ParamByName("HSOLICITAUX").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    vTotalBeneficios = SQL.FieldByName("TOTALBENEFICIOS").AsFloat

    If vPacoteLimiteTipo = "V" Then
      If SQL.FieldByName("TOTALBENEFICIOS").AsFloat >vPacoteLimiteValor Then

        SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio negado. Total dos benefícios excede o limite de adiantamento definitivo do Contrato"

        SQLUPD.ExecSQL

        Set SQLUPD = Nothing

		BsShowMessage("Auxílio Negado. Total dos benefícios excede o limite de valor do pacote", "I")
        SQL.Active = False
        Set SQL = Nothing
        RefreshNodesWithTable("SAM_SOLICITAUX")
        Exit Sub
      End If
    End If

    If SQL.FieldByName("TOTALBENEFICIOS").AsFloat >vLimiteAdiantamentoDefinitivo Then

      SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio negado. Total dos benefícios excede o limite de adiantamento definitivo do Contrato"

      SQLUPD.ExecSQL

      Set SQLUPD = Nothing

	  BsShowMessage("Auxílio Negado. Total dos benefícios excede o limite de adiantamento definitivo do Contrato", "I")
      SQL.Active = False
      Set SQL = Nothing
      RefreshNodesWithTable("SAM_SOLICITAUX")
      Exit Sub
    End If

  End If

End If

Set SQL = Nothing

Set SQLUPD = NewQuery

SQLUPD.Clear
SQLUPD.Add("UPDATE SAM_SOLICITAUX SET")
SQLUPD.Add("  OCORRENCIAS = :OCORRENCIAS,")
SQLUPD.Add("  SITUACAO = 'L'")
SQLUPD.Add("WHERE HANDLE = :HSOLICITAUX")
SQLUPD.ParamByName("HSOLICITAUX").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio liberado "
SQLUPD.ExecSQL

If InTransaction Then Commit

Set SQLUPD = Nothing

RefreshNodesWithTable("SAM_SOLICITAUX")
End Sub

Public Sub BOTAONEGAR_OnClick()
  'SMS 82885 - Artur - Bloqueio de botões quando o registro está em modo de edição
  Dim SQLUPD As Object
  If CurrentQuery.State > 1 Then
	BsShowMessage("O registro está em edição", "I")
  	Exit Sub
  End If

  Dim SQLUSER As Object
  Set SQLUSER = NewQuery

  SQLUSER.Active = False
  SQLUSER.Clear
  SQLUSER.Add("SELECT NOME FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HUSUARIO")
  SQLUSER.ParamByName("HUSUARIO").AsInteger = CurrentUser
  SQLUSER.Active = True
  If CurrentQuery.FieldByName("SITUACAO").AsString = "N" Then
    BsShowMessage("Solicitação já negada", "I")
    Exit Sub
  End If
  'Só nega se não estiver cancelada
  If CurrentQuery.FieldByName("SITUACAO").AsString <>"C" Then
    If (BsShowMessage("Deseja Negar?", "Q") = vbYes) Then

    Set SQLUPD = NewQuery

    If Not InTransaction Then StartTransaction

    SQLUPD.Clear
    SQLUPD.Add("UPDATE SAM_SOLICITAUX SET")
    SQLUPD.Add("  OCORRENCIAS = :OCORRENCIAS,")
    SQLUPD.Add("  SITUACAO = 'N'")
    SQLUPD.Add("WHERE HANDLE = :HSOLICITAUX")
    SQLUPD.ParamByName("HSOLICITAUX").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio negado"
    SQLUPD.ExecSQL

    If InTransaction Then Commit

    Set SQLUPD = Nothing

    SQLUSER.Active = False
    Set SQLUSER = Nothing
    RefreshNodesWithTable("SAM_SOLICITAUX")
  End If
Else
  BsShowMessage("Não é possível negar solicitação cancelada", "I")
End If
End Sub

Public Sub BOTAOPACOTE_OnClick()
  Dim SQLUPD As Object
  Dim interface As Object
  If CurrentQuery.State = 1 Then
    If CurrentQuery.FieldByName("SITUACAO").AsString <>"A" Then
      BsShowMessage("Pacote auxílio permitido somente para solicitações abertas", "I")
      Exit Sub
    End If
    Set interface = CreateBennerObject("SamSolicitAux.Rotinas")
    Set interface.PacoteAuxilio(CurrentSystem, CurrentQuery.FieldByName("PACOTEAUXILIO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    'RefreshNodesWithTable("SAM_SOLICITAUX")
  Else
    BsShowMessage("O registro não pode estar em edição", "I")
  End If
  Set interface = Nothing
End Sub

Public Sub BOTAOTRASFERIR_OnClick()
  'SMS 82885 - Artur - Bloqueio de botões quando o registro está em modo de edição
  Dim SQLUPD As Object
  If CurrentQuery.State > 1 Then
  	BsShowMessage("O registro está em edição", "I")
  	Exit Sub
  End If

  Dim SQLUSER As Object
  Set SQLUSER = NewQuery

  SQLUSER.Active = False
  SQLUSER.Clear
  SQLUSER.Add("SELECT NOME FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HUSUARIO")
  SQLUSER.ParamByName("HUSUARIO").AsInteger = CurrentUser
  SQLUSER.Active = True
  If CurrentQuery.FieldByName("SITUACAO").AsString = "T" Then
  	BsShowMessage("Solicitação já transferida", "I")
    Exit Sub
  End If
  'Só trasferi se for Aberta ou Liberada
  If((CurrentQuery.FieldByName("SITUACAO").AsString = "A")Or(CurrentQuery.FieldByName("SITUACAO").AsString = "L"))Then
  If BsShowMessage("Deseja transferir?", "Q") = vbYes Then

    Set SQLUPD = NewQuery

    If Not InTransaction Then StartTransaction

    'André - 15/10/2004
    Dim Obj As Object
    Set Obj = CreateBennerObject("SAMSOLICITAUX.ExportaSolicitAux")
    If Obj.Executar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) = True Then

      SQLUPD.Clear
      SQLUPD.Add("UPDATE SAM_SOLICITAUX SET")
      SQLUPD.Add("  OCORRENCIAS = :OCORRENCIAS,")
      SQLUPD.Add("  SITUACAO = 'T'")
      SQLUPD.Add("WHERE HANDLE = :HSOLICITAUX")
      SQLUPD.ParamByName("HSOLICITAUX").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQLUPD.ParamByName("OCORRENCIAS").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Str(ServerNow) + "  " + SQLUSER.FieldByName("NOME").AsString + Chr(13) + "Auxílio transferido"
      SQLUPD.ExecSQL

      If InTransaction Then Commit

      Set SQLUPD = Nothing

    End If
    RefreshNodesWithTable("SAM_SOLICITAUX")
  End If
Else
  BsShowMessage("Não é possível trasferir", "I")
End If
End Sub

Public Sub CID_OnEnter()
  CID.AnyLevel = True
End Sub

Public Sub PACOTEAUXILIO_OnChange()
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT PERCEMPRESA, PERCAUXILIO, PERCADIANTAMENTO")
  SQL.Add("FROM SAM_PACOTEAUXILIO")
  SQL.Add("WHERE HANDLE = :HPACOTE")
  SQL.ParamByName("HPACOTE").Value = CurrentQuery.FieldByName("PACOTEAUXILIO").AsInteger
  SQL.Active = True

  CurrentQuery.FieldByName("PERCEMPRESA").Value = SQL.FieldByName("PERCEMPRESA").AsFloat
  CurrentQuery.FieldByName("PERCAUXILIO").Value = SQL.FieldByName("PERCAUXILIO").AsFloat
  CurrentQuery.FieldByName("PERCADIANTAMENTO").Value = SQL.FieldByName("PERCADIANTAMENTO").AsFloat

  Set SQL = Nothing

End Sub

Public Sub TABLE_AfterInsert()
  Dim SQL As Object
  Set SQL = NewQuery

  'SQL p/gerar novo codigo
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT MAX(CODIGO) AS MAXIMO FROM SAM_SOLICITAUX")
  SQL.Active = True
  CurrentQuery.FieldByName("CODIGO").AsInteger = SQL.FieldByName("MAXIMO").AsInteger + 1
  ROTULOBEN01.Text = ""
  ROTULOBEN02.Text = ""

  SQL.Active = False
  Set SQL = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQLBEN As Object
  Dim SQLFAM As Object
  Dim SQLTIT As Object
  Dim SQLCON As Object
  Set SQLCON = NewQuery
  Set SQLBEN = NewQuery
  Set SQLFAM = NewQuery
  Set SQLTIT = NewQuery

  SQLBEN.Active = False
  SQLBEN.Clear
  SQLBEN.Add("SELECT NOME,FAMILIA,CONTRATO FROM SAM_BENEFICIARIO")
  SQLBEN.Add("WHERE HANDLE = :CBENEFICIARIO")
  SQLBEN.ParamByName("CBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQLBEN.Active = True

  SQLCON.Active = False
  SQLCON.Clear
  SQLCON.Add("SELECT CONTRATO,CNPJ,CPFRESPONSAVEL FROM SAM_CONTRATO WHERE HANDLE = :HCONTRATO")
  SQLCON.ParamByName("HCONTRATO").AsInteger = SQLBEN.FieldByName("CONTRATO").AsInteger
  SQLCON.Active = True

  SQLFAM.Active = False
  SQLFAM.Clear
  SQLFAM.Add("SELECT TITULARRESPONSAVEL,TABRESPONSAVEL FROM SAM_FAMILIA WHERE HANDLE = :HFAMILIA")
  SQLFAM.ParamByName("HFAMILIA").AsInteger = SQLBEN.FieldByName("FAMILIA").AsInteger
  SQLFAM.Active = True

  SQLTIT.Active = False
  SQLTIT.Clear
  SQLTIT.Add("SELECT NOME FROM SAM_BENEFICIARIO WHERE HANDLE = :HBENEFICIARIO")
  SQLTIT.ParamByName("HBENEFICIARIO").AsInteger = SQLFAM.FieldByName("TITULARRESPONSAVEL").AsInteger
  SQLTIT.Active = True

  If SQLBEN.EOF And CurrentQuery.State = 1 Then
    ROTULOBEN01.Text = "*** BENEFICIARIO NÃO ENCONTRADO ***"
    ROTULOBEN02.Text = "*"
  Else
    If SQLFAM.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
      ROTULOBEN01.Text = "Beneficiário : " + SQLBEN.FieldByName("NOME").AsString
      ROTULOBEN02.Text = "Titular Responsável - Beneficiário: " + Format(SQLCON.FieldByName("CPFRESPONSAVEL").AsString, "000\.000\.000\-00") + " - " + SQLTIT.FieldByName("NOME").AsString
    Else
      ROTULOBEN01.Text = "Beneficiário : " + SQLBEN.FieldByName("NOME").AsString
      ROTULOBEN02.Text = "Titular Responsável - Pessoa: " + Format(SQLCON.FieldByName("CNPJ").AsString, "00\.000\.000\/0000\-00") + " - " + SQLTIT.FieldByName("NOME").AsString
    End If
  End If

  SQLBEN.Active = False
  SQLFAM.Active = False
  SQLTIT.Active = False
  SQLCON.Active = False
  Set SQLCON = Nothing
  Set SQLBEN = Nothing
  Set SQLFAM = Nothing
  Set SQLTIT = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object
  Set SQL = NewQuery
  Dim vTotal As Double

  'Balani SMS 54421 10/01/2006
  If CurrentQuery.FieldByName("DATAVALIDADE").AsDateTime < CurrentQuery.FieldByName("DATACADASTRO").AsDateTime Then
    BsShowMessage("Data de validade menor que a data de cadastro.", "I")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If
  'final SMS 54421

  If(CurrentQuery.State = 3)And _
     (CurrentQuery.FieldByName("TABCLASSIFICACAO").AsInteger = 2)Then
  SQL.Clear
  SQL.Add("SELECT HANDLE")
  SQL.Add("FROM SAM_SOLICITAUX")
  SQL.Add("WHERE BENEFICIARIO = :HBENEFICIARIO")
  SQL.Add("  AND TABCLASSIFICACAO = 2")
  SQL.Add("  AND SITUACAO = 'A'")
  SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    BsShowMessage("Existe uma solicitação de deslocamento aberta para este beneficiário", "I")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If
End If

'Verifica se o campo NAOPERMITEAUXILIO está marcado
SQL.Clear
SQL.Active = False
SQL.Add("SELECT NAOPERMITEAUXILIO, NOME FROM SAM_BENEFICIARIO WHERE HANDLE = :HBeneficiario ")
SQL.ParamByName("HBeneficiario").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
SQL.Active = True

If SQL.FieldByName("NAOPERMITEAUXILIO").AsString = "S" Then
  BsShowMessage("O beneficiário : " + SQL.FieldByName("NOME").AsString + " não permite auxílio", "I")
  SQL.Active = False
  Set SQL = Nothing
  CanContinue = False
  Exit Sub
End If

If Not DVOk Then
  BsShowMessage("DV Inválido", "I")
  CanContinue = False
  DV.SetFocus
  Exit Sub
End If

If Not CurrentQuery.FieldByName("PACOTEAUXILIO").IsNull Then
  SQL.Clear
  'Balani SMS 54421 09/01/2006
  SQL.Add("SELECT P.HANDLE")
  SQL.Add("  FROM SAM_PACOTEAUXILIO P")
  SQL.Add("  JOIN SAM_BENEFICIARIO B ON (B.HANDLE = :BENEFICIARIO)")
  SQL.Add("  JOIN SAM_CONTRATO_AUXILIO CA ON (CA.CONTRATO = B.CONTRATO)")
  SQL.Add("  JOIN SAM_CONTRATO_AUXILIOPACOTE CAP ON (CAP.CONTRATOAUXILIO = CA.HANDLE)")
  SQL.Add(" WHERE P.HANDLE = :HPACOTEAUXILIO")
  SQL.Add("   AND CAP.PACOTEAUXILIO = P.HANDLE")
  SQL.Add("   AND CA.DATAINICIAL <= B.DATAADMISSAO")
  SQL.Add("   AND (CA.DATAFINAL IS NULL OR CA.DATAFINAL >= B.DATAADMISSAO)")
  SQL.Add("   AND CAP.DATAINICIAL <= :DATACADASTRO")
  SQL.Add("   AND (CAP.DATAFINAL IS NULL OR CAP.DATAFINAL >= :DATAVALIDADE)")
  SQL.ParamByName("HPACOTEAUXILIO").Value = CurrentQuery.FieldByName("PACOTEAUXILIO").AsInteger
  SQL.ParamByName("DATACADASTRO").Value = CurrentQuery.FieldByName("DATACADASTRO").AsDateTime
  SQL.ParamByName("DATAVALIDADE").Value = CurrentQuery.FieldByName("DATAVALIDADE").AsDateTime
  SQL.ParamByName("BENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger

  'final SMS 54421
  SQL.Active = True

  If SQL.EOF Then
    CanContinue = False
    BsShowMessage("Data de cadastro fora do período de vigência do pacote, ou beneficiário não tem direito de utilização ao pacote de auxílio.", "I")
    Set SQL = Nothing
    Exit Sub
  End If
End If


If CurrentQuery.FieldByName("TABCLASSIFICACAO").AsInteger = 1 Then
'Balani SMS 49282 06/10/2005
  vTotal = CurrentQuery.FieldByName("PERCEMPRESA").AsFloat + CurrentQuery.FieldByName("PERCAUXILIO").AsFloat + CurrentQuery.FieldByName("PERCADIANTAMENTO").AsFloat
  'If(CurrentQuery.FieldByName("PERCEMPRESA").AsFloat + CurrentQuery.FieldByName("PERCAUXILIO").AsFloat + CurrentQuery.FieldByName("PERCADIANTAMENTO").AsFloat)<>100 Then
  If (vTotal < 0.01) Or (vTotal > 100) Then
    CanContinue = False
    'MsgBox("A soma dos percentuais de Empresa, Auxílio e de Adiantamento deve ser igual a 100%")
    BsShowMessage("A soma dos percentuais de Empresa, Auxílio e Adiantamento devem estar entre 0.01% e 100%", "I")
'Final SMS 49282
    Exit Sub
  End If
End If

SQL.Active = False
Set SQL = Nothing
End Sub

Public Function DVOk As Boolean
  Dim OLEAutorizador As Object
  DVOk = True
  If Not CurrentQuery.FieldByName("DV").IsNull Then
    Set OLEAutorizador = CreateBennerObject("SamAuto.Autorizador")
    On Error GoTo erro
    OLEAutorizador.Inicializar(CurrentSystem)
    DVOk = OLEAutorizador.BeneficiarioDvOk(CurrentSystem, CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, CurrentQuery.FieldByName("DV").AsString, CurrentQuery.FieldByName("DATACADASTRO").AsDateTime)
    Set OLEAutorizador = Nothing
  End If
  Exit Function
erro :
  Set OLEAutorizador = Nothing
  BsShowMessage("(AUT002)Erro interno, contate revendedor", "I")
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOLIBERAR"
			BOTAOLIBERAR_OnClick
		Case "BOTAONEGAR"
			BOTAONEGAR_OnClick
		Case "BOTAOTRANSFERIR"
			BOTAOTRASFERIR_OnClick
	End Select
End Sub
