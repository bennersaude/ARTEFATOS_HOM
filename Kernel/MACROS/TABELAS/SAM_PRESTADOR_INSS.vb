'HASH: 5FB230EF9376CD1F690BD185FCF385FA
'Macro: SAM_PRESTADOR_INSS

'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()

  Dim QryInss As Object
  Set QryInss = NewQuery

  QryInss.Active = False
  QryInss.Clear
  QryInss.Add("SELECT INSCRICAOINSS     ")
  QryInss.Add("  FROM SAM_PRESTADOR     ")
  QryInss.Add(" WHERE HANDLE = :PHANDLE ")
  QryInss.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  'RecordHandleOfTable("SAM_PRESTADOR")
  QryInss.Active = True

  If Not QryInss.FieldByName("INSCRICAOINSS").IsNull Then
    CurrentQuery.FieldByName("INSCRICAO").AsString = QryInss.FieldByName("INSCRICAOINSS").AsString
    INSCRICAO.ReadOnly = True
  Else
    INSCRICAO.ReadOnly = False
  End If

  QryInss.Active = False

  Set QryInss = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  '
  If (Len(CurrentQuery.FieldByName("CODIGOPAGTOGPS").AsString) > 0) And (Len(CurrentQuery.FieldByName("CODIGOPAGTOGPS").AsString) < 4) Then
    bsShowMessage("Código de pagamento do GPS, deve ter obrigatoriamente 4 dígitos!", "E")
    CanContinue = False
    Exit Sub
  End If

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

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 1 Then
    CurrentQuery.FieldByName("CNPJEMPREGADOR").Clear
    CurrentQuery.FieldByName("EMPREGADOR").Clear
  End If

  If Not CurrentQuery.FieldByName("CNPJEMPREGADOR").IsNull Then
    If Not IsValidCGC(CurrentQuery.FieldByName("CNPJEMPREGADOR").AsString) Then
      CanContinue = False
      bsShowMessage("CNPJ Inválido!", "E")
      Exit Sub
    End If
  End If


  'Dim P As Object
  'Set P = NewQuery
  'P.Clear
  'P.Add("SELECT INSSFAIXAS FROM SFN_PARAMETROSFIN")
  'P.Active = True

  'If CurrentQuery.FieldByName("FAIXA").AsInteger > P.FieldByName("INSSFAIXAS").AsInteger Then
  '    CanContinue = False
  '    MsgBox "Número de faixa inválido! Válido somente até "+Str( P.FieldByName("INSSFAIXAS").AsInteger)
  '    Exit Sub
  'End If
  'P.Active = False

  Set P = Nothing

  If VisibleMode = True Or WebMode Then
    'CHECA VIGENCIA
    Dim Interface As Object
    Dim Linha As String

    Linha = ""

    Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

    Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_INSS", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "PRESTADOR", "")

    If Linha = "" Then
      CanContinue = True
    Else
      CanContinue = False
      bsShowMessage(Linha, "E")
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("INSCRICAO").AsString = "" Then
    Dim QryInss As Object
    Set QryInss = NewQuery

    CancelDescription = "BUSCA INSCRICAO"

    QryInss.Active = False
    QryInss.Clear
    QryInss.Add("SELECT INSCRICAOINSS     ")
    QryInss.Add("  FROM SAM_PRESTADOR     ")
    QryInss.Add(" WHERE HANDLE = :PHANDLE ")
    QryInss.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
    'RecordHandleOfTable("SAM_PRESTADOR")
    QryInss.Active = True

    If Not QryInss.FieldByName("INSCRICAOINSS").IsNull Then
      CurrentQuery.FieldByName("INSCRICAO").AsString = QryInss.FieldByName("INSCRICAOINSS").AsString
      INSCRICAO.ReadOnly = True
    Else
      INSCRICAO.ReadOnly = False
    End If

    QryInss.Active = False

    Set QryInss = Nothing
  Else 'SMS 87115 - Marcelo Barbosa - 11/09/2007
    If Not IsNumeric(CurrentQuery.FieldByName("INSCRICAO").AsString) Then
     bsShowMessage("Número de Inscrição não pode conter letras. Somente Números e ponto.", "I")
     CanContinue = False
    End If
  End If

End Sub

Public Sub xxTABTIPO_OnChanging(AllowChange As Boolean)
  If VisibleMode = True Then
    MsgBox("Alteração não permitida")
    AllowChange = False
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

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

