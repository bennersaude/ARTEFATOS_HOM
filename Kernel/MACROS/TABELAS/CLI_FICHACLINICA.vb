'HASH: AE61B473CCA8D5A20A2F1B21A857F906

'CLI_FICHACLINICA

Public Sub TABLE_AfterScroll()
  Dim NOME As String
  Dim SEXO As String
  Dim IDADE As String
  Dim DATAADESAO As String
  Dim LOTACAO As String
  Dim ESTADOCIVIL As String
  Dim NATURALIDADE As String
  Dim LOGRESIDENCIAL As String
  Dim NUMRESIDENCIAL As String
  Dim MUNRESIDENCIAL As String
  Dim LOGCOMERCIAL As String
  Dim NUMCOMERCIAL As String
  Dim MUNCOMERCIAL As String
  Dim CONTRATO As String
  Dim ANOS As Long
  Dim MESES As Long
  Dim DIAS As Long
  Dim SQL As Object
  Set SQL = NewQuery
  Dim PRONTUARIO As Object
  Set PRONTUARIO = CreateBennerObject("CliProntuario.Rotinas")
  ROTULONOME.Text = ""

  SQL.Clear
  SQL.Add("SELECT M.NOME,")
  SQL.Add("       M.SEXO,")
  SQL.Add("       M.DATANASCIMENTO,")
  SQL.Add("       B.DATAADESAO,")
  SQL.Add("       CL.DESCRICAO LOTACAO,")
  SQL.Add("       ES.DESCRICAO ESTADOCIVIL,")
  SQL.Add("       MU.NOME NATURALIDADE,")
  SQL.Add("       EC.LOGRADOURO LOGCOMERCIAL,")
  SQL.Add("       EC.NUMERO NUMCOMERCIAL,")
  SQL.Add("       MC.NOME MUNCOMERCIAL,")
  SQL.Add("       ER.LOGRADOURO LOGRESIDENCIAL,")
  SQL.Add("       ER.NUMERO NUMRESIDENCIAL,")
  SQL.Add("       MR.NOME MUNRESIDENCIAL,")
  SQL.Add("       C.CONTRATANTE CONTRATO")

    SQL.Add("  FROM CLI_ATENDIMENTO A")
    SQL.Add("  LEFT JOIN SAM_MATRICULA M ON (A.MATRICULA = M.HANDLE)")
    SQL.Add("  LEFT JOIN SAM_BENEFICIARIO B ON (B.MATRICULA = M.HANDLE)")
    SQL.Add("  LEFT JOIN SAM_FAMILIA F ON (B.FAMILIA = F.HANDLE)")
    SQL.Add("  LEFT JOIN SAM_ESTADOCIVIL ES  ON (B.ESTADOCIVIL = ES.HANDLE)")
    SQL.Add("  LEFT JOIN MUNICIPIOS MU ON (M.MUNICIPIO = MU.HANDLE)")
    SQL.Add("  LEFT JOIN SAM_ENDERECO EC ON (B.ENDERECOCOMERCIAL = EC.HANDLE)")
    SQL.Add("  LEFT JOIN MUNICIPIOS MC ON (EC.MUNICIPIO = MC.HANDLE)")
    SQL.Add("  LEFT JOIN SAM_ENDERECO ER ON (B.ENDERECORESIDENCIAL = ER.HANDLE)")
    SQL.Add("  LEFT JOIN MUNICIPIOS MR ON (ER.MUNICIPIO = MR.HANDLE)")
    SQL.Add("  LEFT JOIN SAM_CONTRATO C ON (B.CONTRATO = C.HANDLE)")
    SQL.Add("  LEFT JOIN SAM_CONTRATO_LOTACAO CL ON (F.LOTACAO = CL.HANDLE)")
    SQL.Add(" WHERE A.HANDLE = :ATENDIMENTO")

  SQL.ParamByName("ATENDIMENTO").AsInteger = RecordHandleOfTable("CLI_ATENDIMENTO")
  '-----------------------------------------
  
  '-----------------------------------------
  SQL.Active = True

  If Not SQL.FieldByName("DATANASCIMENTO").IsNull Then
    PRONTUARIO.Idade(CurrentSystem, SQL.FieldByName("DATANASCIMENTO").AsDateTime, DIAS, MESES, ANOS)
  End If
  NOME = SQL.FieldByName("NOME").AsString
  SEXO = SQL.FieldByName("SEXO").AsString
  If SEXO <>"" Then
    If SEXO = "F" Then
      SEXO = "Feminino"
    Else
      SEXO = "Masculino"
    End If
  End If
  If ANOS >0 Then
    IDADE = Str(ANOS) + " ano(s), " + Str(MESES) + " mese(s) e " + Str(DIAS) + " dia(s)"
  End If
  DATAADESAO = SQL.FieldByName("DATAADESAO").AsString
  LOTACAO = SQL.FieldByName("LOTACAO").AsString
  ESTADOCIVIL = SQL.FieldByName("ESTADOCIVIL").AsString
  NATURALIDADE = SQL.FieldByName("NATURALIDADE").AsString
  LOGCOMERCIAL = SQL.FieldByName("LOGCOMERCIAL").AsString
  NUMCOMERCIAL = SQL.FieldByName("NUMCOMERCIAL").AsString
  MUNCOMERCIAL = SQL.FieldByName("MUNCOMERCIAL").AsString
  If NUMCOMERCIAL <>"" Then
    LOGCOMERCIAL = LOGCOMERCIAL + ", " + NUMCOMERCIAL
  End If
  If MUNCOMERCIAL <>"" Then
    LOGCOMERCIAL = LOGCOMERCIAL + " - " + MUNCOMERCIAL
  End If
  LOGRESIDENCIAL = SQL.FieldByName("LOGRESIDENCIAL").AsString
  NUMRESIDENCIAL = SQL.FieldByName("NUMRESIDENCIAL").AsString
  MUNRESIDENCIAL = SQL.FieldByName("MUNRESIDENCIAL").AsString
  If NUMRESIDENCIAL <>"" Then
    LOGRESIDENCIAL = LOGRESIDENCIAL + ", " + NUMRESIDENCIAL
  End If
  If MUNRESIDENCIAL <>"" Then
    LOGRESIDENCIAL = LOGRESIDENCIAL + " - " + MUNRESIDENCIAL
  End If
  CONTRATO = SQL.FieldByName("CONTRATO").AsString
  If NOME <>"" Then
    ROTULONOME.Text = "Nome: " + NOME + Chr(13)
  End If
  If SEXO <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Sexo: " + SEXO + Chr(13)
  End If
  If IDADE <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Idade: " + IDADE + Chr(13)
  End If
  If DATAADESAO <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Data de adesão: " + DATAADESAO + Chr(13)
  End If
  If LOTACAO <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Lotação: " + LOTACAO + Chr(13)
  End If
  If ESTADOCIVIL <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Estado civil: " + ESTADOCIVIL + Chr(13)
  End If
  If NATURALIDADE <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Naturalidade: " + NATURALIDADE + Chr(13)
  End If
  If LOGRESIDENCIAL <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Endereço residencial: " + LOGRESIDENCIAL + Chr(13)
  End If
  If LOGCOMERCIAL <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Endereço comercial: " + LOGCOMERCIAL + Chr(13)
  End If
  If CONTRATO <>"" Then
    ROTULONOME.Text = ROTULONOME.Text + "Contrato: " + CONTRATO + Chr(13)
  End If
  Set PRONTUARIO = Nothing
  Set SQL = Nothing
End Sub

