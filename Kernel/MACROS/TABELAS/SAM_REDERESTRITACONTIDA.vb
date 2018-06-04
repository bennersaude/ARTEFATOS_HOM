'HASH: FF0EC353F44A6024BC1A681DEDF215F9

'Macro Tabela:  SAM_REDERESTRITACONTIDA
'#Uses "*bsShowMessage"

Dim res As Boolean
Dim Rede As Long
Dim vgNomePrest As String '--claudemir - 21/10/2002


Public Sub Recursividade(pREDERESTRITA As Long)
  Dim CONTIDAS As Object

  If pREDERESTRITA = Rede Then
    res = False
  Else
    Set CONTIDAS = NewQuery

    CONTIDAS.Add("SELECT REDERESTRITACONTIDA FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITA = :REDE")
    CONTIDAS.ParamByName("REDE").Value = pREDERESTRITA
    CONTIDAS.Active = True
    While (Not CONTIDAS.EOF) And (res)
      Recursividade(CONTIDAS.FieldByName("REDERESTRITACONTIDA").Value)
      CONTIDAS.Next
    Wend
  End If
End Sub



Public Sub TABLE_AfterPost()
  Dim Interface As Object
  Dim vOperacao As String

  'registro só edita incluindo e via tree-view só inclusão --- exclusão só via interface.
  If CurrentQuery.State = 1 Then
    vOperacao = "I"
  End If

  Set Interface = CreateBennerObject("BSPRE001.Rotinas")
  Interface.AtualizaContidas(CurrentSystem, CurrentQuery.FieldByName("REDERESTRITA").AsInteger, CurrentQuery.FieldByName("REDERESTRITACONTIDA").AsInteger, vOperacao)
  Set Interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


  Dim SQL1 As Object
  Dim SQL2 As Object
  Dim SQL3 As Object


  If CurrentQuery.FieldByName("REDERESTRITACONTIDA").Value = CurrentQuery.FieldByName("REDERESTRITA").Value Then
    CanContinue = False
    bsShowMessage("A rede contida deve ser diferente da rede acima !!!", "E")
  Else
    Set SQL2 = NewQuery
    SQL2.Add("SELECT HANDLE FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITA = :REDERESTRITACONTIDA AND REDERESTRITACONTIDA = :REDERESTRITA")
    SQL2.ParamByName("REDERESTRITA").Value = CurrentQuery.FieldByName("REDERESTRITA").Value
    SQL2.ParamByName("REDERESTRITACONTIDA").Value = CurrentQuery.FieldByName("REDERESTRITACONTIDA").Value
    SQL2.Active = True
    If Not SQL2.FieldByName("HANDLE").IsNull Then
      CanContinue = False
      bsShowMessage("Existe uma composição invertida !!!", "E")
    End If
  End If

  Set SQL1 = NewQuery

  SQL1.Add("SELECT HANDLE FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITA = :REDERESTRITA AND REDERESTRITACONTIDA = :REDERESTRITACONTIDA")
  SQL1.ParamByName("REDERESTRITA").Value = CurrentQuery.FieldByName("REDERESTRITA").Value
  SQL1.ParamByName("REDERESTRITACONTIDA").Value = CurrentQuery.FieldByName("REDERESTRITACONTIDA").Value
  SQL1.Active = True

  If Not SQL1.FieldByName("HANDLE").IsNull Then
    CanContinue = False
    bsShowMessage("Está composição já existe !!!", "E")
  End If

  Set SQL1 = Nothing

  Set SQL3 = NewQuery
  SQL3.Add("SELECT REDERESTRITA, REDERESTRITACONTIDA FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITA = :REDERESTRITACONTIDA")
  SQL3.ParamByName("REDERESTRITACONTIDA").Value = CurrentQuery.FieldByName("REDERESTRITACONTIDA").Value
  SQL3.Active = True

  Rede = CurrentQuery.FieldByName("REDERESTRITA").Value

  res = True
  While (Not SQL3.EOF) And (res)
    Recursividade(SQL3.FieldByName("REDERESTRITACONTIDA").Value)
    SQL3.Next
  Wend
  If res = False Then
    CanContinue = False
    bsShowMessage("Rede contida em outro nó acima desta árvore !!!", "E")
  End If

  Set SQL = Nothing

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  vFiltro = checkPermissaoFilial (CurrentSystem, "E", "P", Msg)
  If vFitro = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  '--- claudemir ---
  Dim linha As String
  CanContinue = False
  linha = "Operação Cancelada !!!" + Chr(10) + _
          "Exclusão de redes contidas somente via interface" + Chr(10) + _
          "(botão 'Cadastrar Redes Contidas' da carga 'Rede Restrita')"
  bsShowMessage(linha, "E")
  Exit Sub

  '-----------------
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
    ' se For alterar os dados de um prestador já cadastrado
    If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
      qAuxiliar.Add("SELECT FILIALPADRAO")
      qAuxiliar.Add("  FROM SAM_PRESTADOR")
      qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
      qAuxiliar.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      qAuxiliar.Active = True
      If qAuxiliar.FieldByName("FILIALPADRAO").IsNull Then
        qAuxiliar.Clear
        qAuxiliar.Add("SELECT FILIALPADRAO")
        qAuxiliar.Add("  FROM Z_GRUPOUSUARIOS")
        qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
        qAuxiliar.ParamByName("HANDLE").Value = CurrentUser
        qAuxiliar.Active = True
      End If
      ' se For cadastrar um novo prestador
    Else
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
              "(SELECT DISTINCT M.HANDLE " + _
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
              "     AND U.ALTERAR = 'S' ) "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode alterar
    vFiltro = ""
    vFiltro = vFiltro + _
              "(SELECT DISTINCT M.HANDLE " + _
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
              "     AND U.ALTERAR = 'S' )"
  End If
  If pServico = "I" Then
    ' Verifica se pode incluir conforme a filial padrao
    vFiltro = vFiltro + _
              "(SELECT DISTINCT M.HANDLE " + _
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
              "     AND U.INCLUIR = 'S' ) "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode incluir
    vFiltro = ""
    vFiltro = vFiltro + _
              "(Select DISTINCT M.HANDLE " + _
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
              "     AND U.INCLUIR = 'S' ) "

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

