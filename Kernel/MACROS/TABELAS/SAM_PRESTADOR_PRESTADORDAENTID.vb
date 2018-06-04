'HASH: 5A04E508C80F3FD47DC7C0C37E4EECC9
'Macro: SAM_PRESTADOR_PRESTADORDAENTID

'#Uses "*bsShowMessage"

'02/01/2001 -Alterado por Paulo Garcia Junior -liberacao para edição do registro atraves dos parametros gerais de prestador

'--Lacerda SMS 19699 -24.10.2003 --------------------

Public Sub ENTIDADE_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCabecs As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String
  Dim vTitulo As String

  If ENTIDADE.PopupCase <>0 Then
    ShowPopup = False
    Set Interface = CreateBennerObject("Procura.Procurar")

    vCabecs = "Código|Prestador|CPFCNPJ"
    vColunas = "PRESTADOR|NOME|CPFCNPJ"
    vCriterio = ""
    vTabela = "SAM_PRESTADOR"
    vTitulo = "Entidade"

    vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("ENTIDADE").AsInteger = vHandle
    End If
    Set Interface = Nothing
  Else
    ShowPopup = True
  End If
End Sub

'----------------------------------------------------

Public Sub PRESTADOR_OnChange()
  BuscaSexo
End Sub


Public Sub TABLE_AfterInsert()

End Sub

Public Sub TABLE_AfterScroll()
  BuscaSexo
  If liberaCorpoClinico <>"" Then
    DATAFINAL.ReadOnly = True
    DATAINICIAL.ReadOnly = True
    ENTIDADE.ReadOnly = True
    ISENTOIRRF.ReadOnly = True
    PRECO.ReadOnly = True
    PRESTADOR.ReadOnly = True
    SEXO.ReadOnly = True
    TEMPORARIO.ReadOnly = True
  Else
    DATAFINAL.ReadOnly = False
    DATAINICIAL.ReadOnly = False
    ENTIDADE.ReadOnly = False
    ISENTOIRRF.ReadOnly = False
    PRECO.ReadOnly = False
    PRESTADOR.ReadOnly = False
    SEXO.ReadOnly = False
    TEMPORARIO.ReadOnly = False
  End If

  If (VisibleMode And NodeInternalCode = 301004) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PRESTADORDA_1508") Then 'ESTA NA CARGA DE MEMBROS DO CORPOCLÍNICO
    ENTIDADE.ReadOnly = True
    ENTIDADE.Visible = True
  End If

  If (VisibleMode And NodeInternalCode = 301005) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PRESTADORDAENTID") Then 'ESTA NA CARGA DE entidades onde é membro do corpo-clínico
    PRESTADOR.ReadOnly = True
    PRESTADOR.Visible = True
    ROTULOSEXO.Text = ""
  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  Msg = liberaCorpoClinico
  If Msg <>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  Msg = liberaCorpoClinico
  If Msg <>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  'Comentado por Paulo Garcia Junior -02.01.2002
  'Dim SQL As Object
  'Set SQL=NewQuery
  'SQL.Add(GenSql("SAM_PRESTADOR","FISICAJURIDICA","","HANDLE = :HANDLE"))
  'SQL.ParamByName("HANDLE").Value=RecordHandleOfTable("SAM_PRESTADOR")
  'SQL.Active=True
  'If SQL.FieldByName("FISICAJURIDICA").Value =1 Then 'Não pode ser incluido para tipo de pessoa física
  ' 	MsgBox "Inclusão não permitidia para tipo de pessoa FÍSICA"
  '	CurrentQuery.Cancel
  '	RefreshNodesWithTable "SAM_PRESTADOR_PRESTADORDAENTID"
  'End If
  'Set SQL=Nothing

  Dim Msg As String
  Msg = liberaCorpoClinico
  If Msg <>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If
End Sub



Public Function liberaCorpoClinico As String
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT EDITAPRESTADORDAENTIDADE FROM SAM_PARAMETROSPRESTADOR")
  SQL.Active = True
  If SQL.FieldByName("EDITAPRESTADORDAENTIDADE").AsString <>"S" Then
    liberaCorpoClinico = "Carga somente para leitura!"
  End If
End Function


Public Sub BuscaSexo
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT SEXO FROM SAM_PRESTADOR WHERE HANDLE = :H")
  SQL.ParamByName("H").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.Active = True
  If SQL.EOF Then
    ROTULOSEXO.Text = ""
  Else
    If SQL.FieldByName("SEXO").AsString = "M" Then
      ROTULOSEXO.Text = "Sexo: Masculino"
    ElseIf SQL.FieldByName("SEXO").AsString = "F" Then
      ROTULOSEXO.Text = "Sexo: Feminino"
    Else
      ROTULOSEXO.Text = ""
    End If
  End If
End Sub

Public Sub TABLE_NewRecord()
  If (VisibleMode And NodeInternalCode = 301004) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PRESTADORDA_1508") Then 'ESTA NA CARGA DE MEMBROS DO CORPOCLÍNICO
    CurrentQuery.FieldByName("ENTIDADE").Value = RecordHandleOfTable("SAM_PRESTADOR")
  End If

  If (VisibleMode And NodeInternalCode = 301005) Or (WebMode And WebVisionCode = "V_SAM_PRESTADOR_PRESTADORDAENTID") Then 'ESTA NA CARGA DE entidades onde é membro do corpo-clínico
    CurrentQuery.FieldByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim Interface As Object
  Dim Linha As String
  Dim CAMPO As String
  Dim CONDICAO As String

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

  If (Not CurrentQuery.FieldByName("MOTIVOAFASTAMENTO").IsNull) And (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    bsShowMessage("Ao informar um motivo de afastamento é necessário informar uma data final", "E")
    CanContinue = False
    Exit Sub
  End If

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  CONDICAO = " AND ENTIDADE = " + CurrentQuery.FieldByName("ENTIDADE").AsString

  Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_PRESTADORDAENTID", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", CONDICAO)

  If Linha = "" Then
    CanContinue = True
  Else
    bsShowMessage(Linha + " Para este Prestador.", "E")
    CanContinue = False
    Exit Sub
  End If
  Set Interface = Nothing

End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  '***************************  Colocada a procura com filtro no cadastro de membros da entidade **************************
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vFiltro As String

  Set Interface = CreateBennerObject("Procura.Procurar")

  ShowPopup = False

  Dim Msg As String

  vFiltro = " AND SAM_PRESTADOR.HANDLE NOT IN (SELECT PRESTADOR FROM SAM_PRESTADOR_PRESTADORDAENTID WHERE ENTIDADE = " + CurrentQuery.FieldByName("ENTIDADE").AsString + ")"


  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.Z_NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"

  'vCriterio ="SAM_PRESTADOR.FISICAJURIDICA = 1 " +vFiltro
  vCriterio = "SAM_PRESTADOR.HANDLE <> " + CStr(RecordHandleOfTable("SAM_PRESTADOR"))

  vCampos = "CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"
' SMS 93620 - Paulo Melo - 29/02/2008 - Colocado LEFT JOIN nas tabelas ESTADOS e MUNICIPIOS, pois o sistema permitia que apenas prestadores com essas informações preenchidas fossem cadastrados como membros de corpo clínico.
  vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|*ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|*MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestador", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  ShowPopup = False
  Set Interface = Nothing
  '************************************ Durval 13/11/2002 SMS 13758 *********************************************************
End Sub

'Coelho,05/08/2005,SMS 48017 - INSERIDO AS FUNÇÕES DE CHECKPERMISSÃOFILIAL ABAIXO.

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
    qAuxiliar.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ENTIDADE").AsInteger
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
    vFiltro = "SELECT DISTINCT M.HANDLE " + _
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
    vFiltro = " SELECT DISTINCT M.HANDLE " + _
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
              "     AND U.ALTERAR = 'S'  "
  End If
  If pServico = "I" Then
    ' Verifica se pode incluir conforme a filial padrao
    vFiltro = vFiltro + _
              " SELECT DISTINCT M.HANDLE " + _
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
              "     AND U.INCLUIR = 'S'   "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode incluir
    vFiltro = ""
    vFiltro = vFiltro + _
              " Select DISTINCT M.HANDLE " + _
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
              "     AND U.INCLUIR = 'S'   "

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
