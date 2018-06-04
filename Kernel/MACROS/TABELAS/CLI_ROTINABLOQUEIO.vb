'HASH: 726E0CB542704EBB57BA2BE19847A84D
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOGERAR_OnClick()
  Dim vDiasSemana As String
  Dim vErro As Boolean
  Dim vOcorrencia As String
  Dim sql As Object
  Dim vTexto As String

  If CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 2 Then 'ROTINA DE INDISPONIBILIDADE

    If(CurrentQuery.State = 2)Or(CurrentQuery.State = 3)Then
    MsgBox("É necessário salvar a rotina antes de gerá-la!")
    Exit Sub
  End If

  Dim BSCLI001 As Object
  Set BSCLI001 = CreateBennerObject("BSCLI001.ROTINAS")

  vDiasSemana = ""

  If CurrentQuery.FieldByName("DOMINGO").AsString = "S" Then vDiasSemana = vDiasSemana + "|Dom|"
  If CurrentQuery.FieldByName("SEGUNDA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Seg|"
  If CurrentQuery.FieldByName("TERCA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Ter|"
  If CurrentQuery.FieldByName("QUARTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Qua|"
  If CurrentQuery.FieldByName("QUINTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Qui|"
  If CurrentQuery.FieldByName("SEXTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Sex|"
  If CurrentQuery.FieldByName("SABADO").AsString = "S" Then vDiasSemana = vDiasSemana + "|Sab|"

  BSCLI001.GeraIndisponibilidades(CurrentSystem, _
                                  CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime, _
                                  CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime, _
                                  vDiasSemana, _
                                  CurrentQuery.FieldByName("CLINICA").AsInteger, _
                                  CurrentQuery.FieldByName("RECURSO").AsInteger, _
                                  CurrentQuery.FieldByName("INDISPONIBILIDADE").AsInteger, _
                                  vErro, _
                                  vOcorrencia)

  Set BSCLI001 = Nothing

  If Not vErro Then
    If Not InTransaction Then StartTransaction
    Set sql = NewQuery

    vTexto = ""
    If vOcorrencia <>"" Then
      vTexto = "Consultas pendentes no período da rotina:" + Chr(13) + Chr(13) + vOcorrencia
    End If

    sql.Clear
    sql.Add("UPDATE CLI_ROTINABLOQUEIO")
    sql.Add("   SET DATAGERACAO = :DATA,")
    sql.Add("       USUARIOGERACAO = :USUARIO,")
    sql.Add("       OCORRENCIA = :OCORRENCIA")
    sql.Add(" WHERE HANDLE = :HANDLE")
    sql.ParamByName("DATA").AsDateTime = ServerDate
    sql.ParamByName("USUARIO").AsInteger = CurrentUser
    sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.ParamByName("OCORRENCIA").AsMemo = vTexto
    sql.ExecSQL

    If InTransaction Then Commit

    Set sql = Nothing

  End If
Else 'ROTINA DE BLOQUEIO
  If CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 1 Then
    If(CurrentQuery.State = 2)Or(CurrentQuery.State = 3)Then
    MsgBox("É necessário salvar a rotina antes de gerá-la!")
    Exit Sub
  End If

  Dim CLICLINICA As Object
  Set CLICLINICA = CreateBennerObject("CLICLINICA.AGENDA")

  vDiasSemana = ""

  If CurrentQuery.FieldByName("DOMINGO").AsString = "S" Then vDiasSemana = vDiasSemana + "|Dom|"
  If CurrentQuery.FieldByName("SEGUNDA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Seg|"
  If CurrentQuery.FieldByName("TERCA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Ter|"
  If CurrentQuery.FieldByName("QUARTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Qua|"
  If CurrentQuery.FieldByName("QUINTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Qui|"
  If CurrentQuery.FieldByName("SEXTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Sex|"
  If CurrentQuery.FieldByName("SABADO").AsString = "S" Then vDiasSemana = vDiasSemana + "|Sab|"

  CLICLINICA.GeraBloqueios(CurrentSystem, CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime, _
                           CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime, _
                           vDiasSemana, _
                           CurrentQuery.FieldByName("CLINICA").AsInteger, _
                           CurrentQuery.FieldByName("RECURSO").AsInteger, _
                           CurrentQuery.FieldByName("ATIVIDADE").AsInteger, _
                           vErro, _
                           vOcorrencia)

  Set CLICLINICA = Nothing

  If Not vErro Then
    If Not InTransaction Then StartTransaction
    Set sql = NewQuery

    vTexto = ""
    If vOcorrencia <>"" Then
      vTexto = "Consultas marcadas no período da rotina:" + Chr(13) + Chr(13) + vOcorrencia
    End If

    sql.Clear
    sql.Add("UPDATE CLI_ROTINABLOQUEIO")
    sql.Add("   SET DATAGERACAO = :DATA,")
    sql.Add("       USUARIOGERACAO = :USUARIO,")
    sql.Add("       OCORRENCIA = :OCORRENCIA")
    sql.Add(" WHERE HANDLE = :HANDLE")
    sql.ParamByName("DATA").AsDateTime = ServerDate
    sql.ParamByName("USUARIO").AsInteger = CurrentUser
    sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.ParamByName("OCORRENCIA").AsMemo = vTexto
    sql.ExecSQL

    If InTransaction Then Commit

    Set sql = Nothing
  End If
  'início sms 38445 - Edilson.Castro - 10/05/2005
Else '********** ROTINA DE HORÁRIO DE PLANTÃO **********
  If CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 3 Then

    If(CurrentQuery.State = 2)Or(CurrentQuery.State = 3)Then
    MsgBox("É necessário salvar a rotina antes de gerá-la!")
    Exit Sub
  End If

  Dim AGENDADLL As Object
  Set AGENDADLL = CreateBennerObject("CLICLINICA.AGENDA")

  vDiasSemana = ""

  If CurrentQuery.FieldByName("DOMINGO").AsString = "S" Then vDiasSemana = vDiasSemana + "|Dom|"
  If CurrentQuery.FieldByName("SEGUNDA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Seg|"
  If CurrentQuery.FieldByName("TERCA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Ter|"
  If CurrentQuery.FieldByName("QUARTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Qua|"
  If CurrentQuery.FieldByName("QUINTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Qui|"
  If CurrentQuery.FieldByName("SEXTA").AsString = "S" Then vDiasSemana = vDiasSemana + "|Sex|"
  If CurrentQuery.FieldByName("SABADO").AsString = "S" Then vDiasSemana = vDiasSemana + "|Sab|"

  'Buscando handle da tabela SAM_PRESTADOR_ESPECIALIDADE, a partir do handle da especialidade retornado pelo CAEdit
  If CurrentQuery.FieldByName("ESPECIALIDADE").IsNull Then
    Dim SQLEspec As Object
    Dim vEspecialidade As Long
    Set SQLEspec = NewQuery
    SQLEspec.Clear
    SQLEspec.Add("SELECT HANDLE")
    SQLEspec.Add("  FROM SAM_PRESTADOR_ESPECIALIDADE")
    SQLEspec.Add(" WHERE PRESTADOR = (SELECT PRESTADOR FROM CLI_RECURSO WHERE HANDLE = :RECURSO)")
    SQLEspec.Add("   AND ESPECIALIDADE = :ESPECIALIDADE")
    SQLEspec.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
    SQLEspec.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
    SQLEspec.Active = True
    vEspecialidade = SQLEspec.FieldByName("HANDLE").AsInteger
    SQLEspec.Active = False
    Set SQLEspec = Nothing
  Else
    vEspecialidade = 0
  End If

  AGENDADLL.GeraHorarioPlantao(CurrentSystem, _
                               CurrentQuery.FieldByName("CLINICA").AsInteger, _
                               CurrentQuery.FieldByName("RECURSO").AsInteger, _
                               vEspecialidade, _
                               CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime, _
                               CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime, _
                               vDiasSemana, _
                               vErro, _
                               vOcorrencia)

  Set AGENDADLL = Nothing

  If vOcorrencia = "Nenhum horário de plantão gerado!" Then
    MsgBox vOcorrencia
    Exit Sub
  End If

  If Not vErro Then
    If Not InTransaction Then StartTransaction
    Set sql = NewQuery

    sql.Clear
    sql.Add("UPDATE CLI_ROTINABLOQUEIO")
    sql.Add("   SET DATAGERACAO = :DATA,")
    sql.Add("       USUARIOGERACAO = :USUARIO,")
    sql.Add("       OCORRENCIA = :OCORRENCIA")
    sql.Add(" WHERE HANDLE = :HANDLE")
    sql.ParamByName("DATA").AsDateTime = ServerDate
    sql.ParamByName("USUARIO").AsInteger = CurrentUser
    sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.ParamByName("OCORRENCIA").AsMemo = vOcorrencia
    sql.ExecSQL

    If InTransaction Then Commit

    Set sql = Nothing
  End If
End If
End If
'fim sms 38445
End If

RefreshNodesWithTable("CLI_ROTINABLOQUEIO")
End Sub

Public Sub ESPECIALIDADE_OnPopup(ShowPopup As Boolean)
  'início sms 38445 - Edilson.Castro - 10/05/2005
  Dim vPrestador As Long
  Dim sql As Object

  'Buscando handle da tabela prestador, a partir do handle do recurso retornado pelo CAEdit
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT PRESTADOR")
  sql.Add("  FROM CLI_RECURSO")
  sql.Add(" WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
  sql.Active = True
  vPrestador = sql.FieldByName("PRESTADOR").AsInteger
  sql.Active = False
  Set sql = Nothing

  'Selecionando apenas as especialidades possíveis para o prestador (recurso) selecionado
  If Not CurrentQuery.FieldByName("RECURSO").IsNull Then
    ESPECIALIDADE.LocalWhere = "SAM_ESPECIALIDADE.HANDLE IN (" & _
                               "SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE WHERE PRESTADOR = " & CStr(vPrestador) & ")"
  Else
    MsgBox("O recurso ainda não foi definido", vbOkOnly)
    ShowPopup = False
    RECURSO.SetFocus
  End If
  'fim sms 38445
End Sub

Public Sub TABLE_AfterInsert()
  Select Case NodeInternalCode
    Case 2610
      CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 1
    Case 2620
      CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 2
    Case 2630
      CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 3
  End Select
End Sub

Public Sub TABLE_AfterScroll()
  If Not CurrentQuery.FieldByName("DATAGERACAO").IsNull Then
    BOTAOGERAR.Enabled = False
  Else
    BOTAOGERAR.Enabled = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime <= CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime Then
    MsgBox("A data final deve ser maior que a data inicial!")
    CanContinue = False
    Exit Sub
  End If


  'SMS 131975 - renato.felipe
  If (CurrentQuery.FieldByName("TABTIPOROTINA").AsInteger = 2) Then 'ROTINA DE INDISPONIBILIDADE

    Dim sql As BPesquisa
    Set sql = NewQuery
    sql.Clear

    sql.Add("SELECT COUNT(1) QTDE ")
    sql.Add("FROM CLI_ROTINABLOQUEIO ")
    sql.Add("WHERE RECURSO = :RECURSO ")
    sql.Add("AND (")
    sql.Add("     ((DATAHORAINICIAL >= :DATAHORAINICIAL) AND (DATAHORAINICIAL < :DATAHORAFINAL)) OR ")
    sql.Add("     ((DATAHORAFINAL >= :DATAHORAINICIAL) AND (DATAHORAFINAL < :DATAHORAFINAL)) OR ")
    sql.Add("     ((DATAHORAINICIAL <= :DATAHORAINICIAL) AND (DATAHORAFINAL >= :DATAHORAFINAL)) ")
    sql.Add("    ) ")
    sql.Add("AND DATAGERACAO IS NOT NULL ")
    sql.Add("AND TABTIPOROTINA = 2 ")

    sql.ParamByName("RECURSO").Value = CurrentQuery.FieldByName("RECURSO").Value
    sql.ParamByName("DATAHORAINICIAL").Value = CurrentQuery.FieldByName("DATAHORAINICIAL").Value
    sql.ParamByName("DATAHORAFINAL").Value = CurrentQuery.FieldByName("DATAHORAFINAL").Value

    sql.Active = True

    If (sql.FieldByName("QTDE").AsInteger > 0) Then
      bsShowMessage("A data inicial ou final se encontra entre a data inicial e final de uma rotina processada para o recurso selecionado.", "I")
  	  Set sql = Nothing
  	  CanContinue = False
  	  Exit Sub
    End If

    Set sql = Nothing

  End If
  'FIM SMS 131975


End Sub

