'HASH: FFA2B4CDF722BA70A0F5960E44782ABB
'Macro: SAM_REDERESTRITA_PRESTADOR
'Reenvio de macro
'Mauricio Ibelli -sms 1725 -Liberar todos os prestadores para serem cadastrados como rede restrita
'#Uses "*bsShowMessage"

Dim vFiltro As String
Dim vgRede As Long
Dim vgHandle As Long



Public Sub BOTAODUPLICAR_OnClick()
  Dim DuplicaRedeRestritaDLL As Object
  Set DuplicaRedeRestritaDLL = CreateBennerObject("SamDupRedeRestrita.SamDupRedeRestrita")
  DuplicaRedeRestritaDLL.Executar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set DuplicaRedeRestritaDLL = Nothing
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  If CurrentQuery.State = 1 Then
    TABLE_BeforeEdit(ShowPopup)
    If ShowPopup = False Then
      Exit Sub
    End If
  End If

  Dim Interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vUsuario As String
  Dim vFilial As String

  vUsuario = Str(CurrentUser)
  vFilial = Str(CurrentBranch)
  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"

  vCriterio = "MUNICIPIOPAGAMENTO IN (" + vFiltro + ")"

  vCampos = "CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"

  vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If

  Set Interface = Nothing

End Sub

Public Sub TABLE_AfterPost()

  Dim SQL As Object
  ' atualizar a data final nas tabelas filhas
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    Set SQL = NewQuery

    'ACOMODACAO
    SQL.Add("UPDATE SAM_PRECOREDE_ACOMODACAO SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
    'PORTE ANESTÉSICO
    SQL.Clear
    SQL.Add("UPDATE SAM_PRECOREDE_AN SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
    'AUXILIAR
    SQL.Clear
    SQL.Add("UPDATE SAM_PRECOREDE_AUX SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
    'DOTACAO
    SQL.Clear
    SQL.Add("UPDATE SAM_PRECOREDE_DOTAC SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
    'EVENTO FAIXA
    SQL.Clear
    SQL.Add("UPDATE SAM_PRECOREDE_FX SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
    'PORTE DE SALA
    SQL.Clear
    SQL.Add("UPDATE SAM_PRECOREDE_SL SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
    'TIPO SERVICO
    SQL.Clear
    SQL.Add("UPDATE SAM_PRECOREDE_TPSERVICO SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
    'REGIME DOTACAO
    SQL.Clear
    SQL.Add("UPDATE SAM_PRECOREDEREGIME_DOTAC SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL

    'REGIME FAIXA DE EVENTOS
    SQL.Clear
    SQL.Add("UPDATE SAM_PRECOREDEREGIME_FX SET DATAFINAL = :DATAFINAL")
    SQL.Add("WHERE REDERESTRITAPRESTADOR = :HANDLE AND DATAFINAL IS NULL")
    SQL.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL

  End If


  Set SQL = Nothing

End Sub


Function VerificaCadastramento(pRede As Long)As String

  Dim Q1, Q2, Q3 As Object

  Set Q1 = NewQuery
  Q1.Add("SELECT R.REDERESTRITA                   ")
  Q1.Add("  FROM SAM_PRESTADOR_ESPEC_REDE    R,   ")
  Q1.Add("       SAM_PRESTADOR_ESPECIALIDADE E    ")
  Q1.Add(" WHERE E.HANDLE = R.PRESTADORESPECIALIDADE AND E.PRESTADOR = :PRESTADOR AND R.REDERESTRITA = :REDE")
  Q1.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
  Q1.ParamByName("REDE").Value = pRede
  Q1.Active = True
  If Not Q1.EOF Then
    VerificaCadastramento = "nas Especialidade"
    Exit Function
  End If
  Set Q2 = NewQuery
  Q2.Add("SELECT R.REDE                            ")
  Q2.Add("  FROM SAM_PRESTADOR_ESPEC_GRP_REDE   R, ")
  Q2.Add("       SAM_PRESTADOR_ESPECIALIDADEGRP G  ")
  Q2.Add(" WHERE G.HANDLE = R.PRESTADORESPECIALIDADEGRUPO AND G.PRESTADOR = :PRESTADOR AND R.REDE = :REDE ")
  Q2.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
  Q2.ParamByName("REDE").Value = pRede
  Q2.Active = True
  If Not Q2.EOF Then
    VerificaCadastramento = "no Grupo de Eventos das Especialidades"
    Exit Function
  End If
  Set Q3 = NewQuery
  Q3.Add("SELECT REDERESTRITA FROM SAM_PRESTADOR_REGRAREDE WHERE PRESTADOR = :PRESTADOR AND REDERESTRITA = :REDE")
  Q3.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
  Q3.ParamByName("REDE").Value = pRede
  Q3.Active = True

  If Not Q3.EOF Then
    VerificaCadastramento = "nas Regras ou Exceções"
    Exit Function
  End If

End Function


Function Recursividade(pRede As Long)As String

  Dim CONTIDAS As Object
  Dim Ok1 As String

  Ok1 = ""
  Ok1 = VerificaCadastramento(pRede)

  Set CONTIDAS = NewQuery

  CONTIDAS.Add("SELECT REDERESTRITA FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITACONTIDA = :REDERESTRITA")
  CONTIDAS.ParamByName("REDERESTRITA").Value = pRede
  CONTIDAS.Active = True

  While(Not CONTIDAS.EOF)And(Ok1 = "")
  Ok1 = Recursividade(CONTIDAS.FieldByName("REDERESTRITA").AsInteger)
  CONTIDAS.Next
Wend

Recursividade = Ok1

Set CONTIDAS = Nothing

End Function


Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebVisionCode = "V_SAM_REDERESTRITA_PRESTADOR_933" Then
			REDERESTRITA.ReadOnly = True
		End If
		If WebVisionCode = "V_SAM_REDERESTRITA_PRESTADOR_535" Then
			PRESTADOR.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "E", "P", Msg)
  If vFitro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  '---claudemir ---
  Dim Ok As String
  Ok = ""
  Ok = Recursividade(CurrentQuery.FieldByName("REDERESTRITA").AsInteger)
  If Ok <>"" Then
    CanContinue = False
    bsShowMessage("Operação Cancelada !!!" + Chr(10) + "Motivo: Esta rede ou suas redes acima está/estão cadastra(s) " + Ok + " deste prestador.", "E")
  End If
  '-----------------
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not VisibleMode Then
    Exit Sub
  End If

  Dim Msg As String
  vFiltro = checkPermissaoFilial(CurrentSystem, "A", "P", Msg)
  If vFiltro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  If Not CurrentQuery.FieldByName("REDERESTRITA").IsNull Then
    vgRede = CurrentQuery.FieldByName("REDERESTRITA").Value
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String

  vFiltro =checkPermissaoFilial(CurrentSystem,"I","P",Msg)
  If vFiltro ="N" Then
    bsShowMessage(Msg, "E")
    CanContinue =False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  'SMS 49152 - Anderson Lonardoni
  'Esta verificação foi tirada do BeforeInsert e colocada no
  'BeforePost para que, no caso de Inserção, já existam valores
  'no CurrentQuery e para funcionar com o Integrator
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  'SMS 49152 - Fim


  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String


  '---claudemir ---
  If vgRede <>CurrentQuery.FieldByName("REDERESTRITA").Value Then
    '---claudemir ---
    Dim Ok As String
    Ok = ""
    Ok = Recursividade(vgRede)
    If Ok <>"" Then
      CanContinue = False
      bsShowMessage("Operação Cancelada !!!" + Chr(10) + "Motivo: Esta rede ou suas redes acima está/estão cadastra(s) " + Ok + " deste prestador.", "E")
      Exit Sub
    End If
    '-----------------
  End If
  '-----------------

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = "AND PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString

  Linha = Interface.Vigencia(CurrentSystem, "SAM_REDERESTRITA_PRESTADOR", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "REDERESTRITA", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
  Set Interface = Nothing

  '------------------------------------------------
  If Linha = "" Then
    Set Interface = CreateBennerObject("BSPRE001.Rotinas")

    vMsg = Interface.ValidaRedePrestador(CurrentSystem, CurrentQuery.FieldByName("PRESTADOR").Value, CurrentQuery.FieldByName("REDERESTRITA").Value, "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)

    If vMsg = "" Then
      CanContinue = True
    Else
      If bsShowMessage(vMsg + (Chr(13)) + " Deseja Continua?", "Q") = vbYes Then
        CanContinue = True
      Else
        CanContinue = False
      End If
    End If
    Set Interface = Nothing
  End If
  '----------------------------------------------------------------------------------------------------------------
  '******************IMPORTANTE******************** O trecho abaixo evita que uma rede seja fechado se houver algum
  '                                                 cadastro dela na especialidade/grupo/regras-excecoes/


  If CanContinue = True Then

    If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then

      Linha = ""

      Set Interface = CreateBennerObject("BSPRE001.Rotinas")
      Linha = Interface.FecharVigenciaRedePrestador(CurrentSystem, CurrentQuery.FieldByName("PRESTADOR").AsInteger, CurrentQuery.FieldByName("REDERESTRITA").AsInteger, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)

      If Linha <>"" Then
        bsShowMessage(Linha, "E")
        CanContinue = False
      End If

    End If

  End If
  '**********************************************FIM DA VERIFICACAO***************************************************
End Sub


'-----------------------------------------------------------------------------------------------------------------------

Public Function checkPermissaoFilial(CurrentSystem, pServico As String, pTabela As String, pMsg As String)As String

  Dim vFiltro, vResultado As String
  Dim qAuxiliar, qPermissoes, SamPrestadorParametro As Object
  Dim qFilialProc As Object

  Dim SamPrestador

  Set qPermissoes = NewQuery
  Set qFilialProc = NewQuery
  Set qAuxiliar = NewQuery
  Set SamPrestadorParametro = NewQuery

  SamPrestadorParametro.Add(" SELECT CONTROLEDEACESSO, BLOQUEIOFILIALPROCESSAMENTO ")
  SamPrestadorParametro.Add(" FROM SAM_PARAMETROSPRESTADOR  ")
  SamPrestadorParametro.Active = True

  If SamPrestadorParametro.FieldByName("CONTROLEDEACESSO").AsString = "N" Then
    checkPermissaoFilial = "(SELECT HANDLE FROM MUNICIPIOS)"
    pMsg = ""
    Exit Function
  End If

  ' começa o controle de acesso
  ' verifica bloqueio filial de processamento

  If SamPrestadorParametro.FieldByName("BLOQUEIOFILIALPROCESSAMENTO").AsString = "S" Then
    qAuxiliar.Clear
    qAuxiliar.Add("SELECT FILIALPADRAO")
    qAuxiliar.Add("  FROM Z_GRUPOUSUARIOS")
    qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentUser
    qAuxiliar.Active = True
    qFilialProc.Active = False
    qFilialProc.Add("Select FILIALPROCESSAMENTO FROM FILIAIS WHERE HANDLE = :HANDLE")
    qFilialProc.ParamByName("HANDLE").Value = qAuxiliar.FieldByName("FILIALPADRAO").Value
    qFilialProc.Active = True
    If qFilialProc.FieldByName("FILIALPROCESSAMENTO").AsInteger = qAuxiliar.FieldByName("FILIALPADRAO").Value Then
      checkPermissaoFilial = "N"
      pMsg = "Permissão negada! Filial padrão do usuário igual a sua filial de processamento."
      Exit Function
    End If
  End If

  qAuxiliar.Active = False
  qAuxiliar.Clear
  If pTabela = "P" Then
    qAuxiliar.Add("SELECT FILIALPADRAO")
    qAuxiliar.Add("  FROM SAM_PRESTADOR")
    qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
    qAuxiliar.Active = True
    If qAuxiliar.FieldByName("FILIALPADRAO").IsNull Then
      qAuxiliar.Clear
      qAuxiliar.Add("SELECT FILIALPADRAO")
      qAuxiliar.Add("  FROM Z_GRUPOUSUARIOS")
      qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
      qAuxiliar.ParamByName("HANDLE").Value = CurrentUser
      qAuxiliar.Active = True
    End If

    qPermissoes.Active = False
    qPermissoes.Clear
    qPermissoes.Add("SELECT X.ALTERAR, ")
    qPermissoes.Add("       X.INCLUIR, ")
    qPermissoes.Add("       X.EXCLUIR, ")
    qPermissoes.Add("       X.FILIAL ")
    qPermissoes.Add("  FROM (SELECT A.ALTERAR ALTERAR, ")
    qPermissoes.Add("               A.INCLUIR INCLUIR, ")
    qPermissoes.Add("               A.EXCLUIR EXCLUIR, ")
    qPermissoes.Add("               A.FILIAL  FILIAL ")
    qPermissoes.Add("          FROM Z_GRUPOUSUARIOS_FILIAIS A ")
    qPermissoes.Add("         WHERE  A.USUARIO = :USUARIO ")
    qPermissoes.Add("           AND  A.FILIAL  = :FILIAL ")
    qPermissoes.Add("        UNION ")
    qPermissoes.Add("        SELECT U.ALTERAR      ALTERAR, ")
    qPermissoes.Add("               U.INCLUIR      INCLUIR, ")
    qPermissoes.Add("               U.EXCLUIR      EXCLUIR, ")
    qPermissoes.Add("               U.FILIALPADRAO FILIAL ")
    qPermissoes.Add("          FROM Z_GRUPOUSUARIOS U ")
    qPermissoes.Add("         WHERE U.HANDLE = :USUARIO ")
    qPermissoes.Add("           AND U.FILIALPADRAO  = :FILIAL) X ")

    qPermissoes.ParamByName("USUARIO").Value = CurrentUser
    qPermissoes.ParamByName("FILIAL").Value = qAuxiliar.FieldByName("FILIALPADRAO").AsInteger
    qPermissoes.Active = True
  End If


  If pServico = "A" Then
    ' Verifica se pode alterar conforme a filial padrao
    vFiltro = vFiltro + _
              "SELECT DISTINCT M.HANDLE " + _
              "  FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "       SAM_REGIAO R, " + _
              "       MUNICIPIOS M " + _
              " WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "   AND R.FILIAL = A.FILIAL " + _
              "   AND M.REGIAO = R.HANDLE " + _
              "   AND A.ALTERAR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.ALTERAR = 'S'  "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode alterar
    vFiltro = ""
    vFiltro = vFiltro + _
              "SELECT DISTINCT M.HANDLE " + _
              "   FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "        MUNICIPIOS M, " + _
              "        SAM_REGIAO R " + _
              "  WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "    AND M.REGIAO = R.HANDLE " + _
              "    AND A.FILIAL = R.FILIAL " + _
              "    AND A.ALTERAR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.ALTERAR = 'S' "
  End If
  If pServico = "I" Then
    ' Verifica se pode incluir conforme a filial padrao
    vFiltro = vFiltro + _
              "SELECT DISTINCT M.HANDLE " + _
              "  FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "       SAM_REGIAO R, " + _
              "       MUNICIPIOS M " + _
              " WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "   AND R.FILIAL = A.FILIAL " + _
              "   AND M.REGIAO = R.HANDLE " + _
              "   AND A.INCLUIR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.INCLUIR = 'S'  "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode incluir
    vFiltro = ""
    vFiltro = vFiltro + _
              "Select DISTINCT M.HANDLE " + _
              "   FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "        MUNICIPIOS M, " + _
              "        SAM_REGIAO R " + _
              "  WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "    AND M.REGIAO = R.HANDLE " + _
              "    AND A.FILIAL = R.FILIAL " + _
              "    AND A.INCLUIR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE " + _
              "    FROM Z_GRUPOUSUARIOS U, " + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.INCLUIR = 'S'  "

  End If

  ' se não estiver cadastrado
  If(qPermissoes.FieldByName("ALTERAR").IsNull)Then
  If pServico = "" Then
    checkPermissaoFilial = ""
    Exit Function
  End If
End If

' se não informou o servico,retorna uma String com os servicos permitidos "LAIE"
If(pServico = "")Then
vResultado = ""
If qPermissoes.FieldByName("ALTERAR").AsString = "S" Then
  vResultado = vResultado + "A"
End If
If qPermissoes.FieldByName("INCLUIR").AsString = "S" Then
  vResultado = vResultado + "I"
End If
If qPermissoes.FieldByName("EXCLUIR").AsString = "S" Then
  vResultado = vResultado + "E"
End If
' se informou o servico,retorna S/N
Else
  Select Case pServico
    Case "A"
      If qPermissoes.FieldByName("ALTERAR").AsString = "S" Then
        vResultado = "S"
        If(Not qAuxiliar.FieldByName("Handle").IsNull)Then
        vResultado = vFiltro
      Else
        vResultado = "N"
        pMsg = "Permissão negada! Usuário não pode alterar."
      End If
    Else
      vResultado = "N"
      pMsg = "Permissão negada! Usuário não pode alterar."
    End If
  Case "I"
    If qPermissoes.FieldByName("INCLUIR").AsString = "S" Then
      vResultado = "S"
      If(Not qAuxiliar.FieldByName("Handle").IsNull)Then
      vResultado = vFiltro
    Else
      vResultado = "N"
      pMsg = "ermissão negada! Usuário não pode incluir."
    End If
  Else
    vResultado = "N"
    pMsg = "Permissão negada! Usuário não pode incluir."
  End If
Case "E"
  If qPermissoes.FieldByName("EXCLUIR").AsString = "S" Then
    vResultado = "S"
  Else
    vResultado = "N"
    pMsg = "Permissão negada! Usuário não pode excluir."
  End If
End Select
End If
checkPermissaoFilial = vResultado
End Function


Public Function BuscarFiliais(CurrentSystem, prFilial As Long, prFilialProcessamento As Long, prMsg As String)As Boolean

  Dim qPermissoes As Object
  Set qPermissoes = NewQuery

  BuscarFiliais = True
  qPermissoes.Active = False
  qPermissoes.Clear
  qPermissoes.Add("SELECT A.HANDLE, A.FILIALPROCESSAMENTO")
  qPermissoes.Add("FROM   Z_GRUPOUSUARIOS U,             ")
  qPermissoes.Add("       FILIAIS A                      ")
  qPermissoes.Add("WHERE  (U.HANDLE = :USUARIO)          ")
  qPermissoes.Add("AND    (A.HANDLE = U.FILIALPADRAO)    ")
  qPermissoes.ParamByName("USUARIO").Value = CurrentUser
  qPermissoes.Active = True

  If qPermissoes.EOF Then
    prMsg = "Problemas Usuario x Filial."
    Exit Function
  End If

  prFilial = qPermissoes.FieldByName("HANDLE").AsInteger
  prFilialProcessamento = qPermissoes.FieldByName("FILIALPROCESSAMENTO").AsInteger
  prMsg = ""
  BuscarFiliais = False
  Set qPermissoes = Nothing
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAODUPLICAR") Then
		BOTAODUPLICAR_OnClick
	End If
End Sub
