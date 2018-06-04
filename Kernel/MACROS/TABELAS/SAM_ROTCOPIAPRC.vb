'HASH: 7E4BFCC1481CF59FC5540B891C510505
Option Explicit

'#Uses "*ProcuraEvento"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*bsShowMessage"

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Interface As Object
  Dim vOk As Integer

  If CurrentQuery.State = 2 Or CurrentQuery.State = 3 Then
    bsShowMessage("Registro em edição !", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("PROCESSADOPOR").IsNull Then
    bsShowMessage("Está rotina já foi processada !", "I")
    Exit Sub
  End If


  Set Interface = CreateBennerObject("BSPRE007.Rotinas")
  vOk = Interface.Exec(CurrentSystem)

  If vOk = 1 Then
    ' Coelho SMS: 133705 estão sendo atualizados dentro da DLL
    'CurrentQuery.Edit
    'CurrentQuery.FieldByName("PROCESSADOPOR").AsInteger = CurrentUser
    'CurrentQuery.FieldByName("PROCESSADOEM").AsDateTime = ServerDate
    'CurrentQuery.Post
    RefreshNodesWithTable("SAM_ROTCOPIAPRC")
  End If

  Set Interface = Nothing

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


Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)

  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vHandle As Long

  ShowPopup = False


  If CurrentQuery.FieldByName("TABELAGENERICA").IsNull Then
    bsShowMessage("Selecionar Tabela Genérica !", "I")
    Exit Sub
  End If

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|Z_DESCRICAO|DESCRICAOABREVIADA|NIVELAUTORIZACAO"

  vCriterio = vCriterio + "SAM_TGE.HANDLE IN (SELECT EVENTO                  "
  vCriterio = vCriterio + "                     FROM SAM_PRECOGENERICO_DOTAC "
  vCriterio = vCriterio + "                    WHERE TABELAPRECO = " + CurrentQuery.FieldByName("TABELAGENERICA").AsString
  vCriterio = vCriterio + "                  )                               "


  vCampos = "Evento|Descrição|Descrição abreviada|Nível"

  vHandle = Interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 2, vCampos, vCriterio, "Tabela Geral de Eventos", True, EVENTOFINAL.Text, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If


  Set Interface = Nothing

End Sub


Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vHandle As Long

  ShowPopup = False


  If CurrentQuery.FieldByName("TABELAGENERICA").IsNull Then
    bsShowMessage("Selecionar Tabela Genérica !", "I")
    Exit Sub
  End If

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|Z_DESCRICAO|DESCRICAOABREVIADA|NIVELAUTORIZACAO"

  vCriterio = vCriterio + "SAM_TGE.HANDLE IN (SELECT EVENTO                  "
  vCriterio = vCriterio + "                     FROM SAM_PRECOGENERICO_DOTAC "
  vCriterio = vCriterio + "                    WHERE TABELAPRECO = " + CurrentQuery.FieldByName("TABELAGENERICA").AsString
  vCriterio = vCriterio + "                  )                               "


  vCampos = "Evento|Descrição|Descrição abreviada|Nível"

  vHandle = Interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 2, vCampos, vCriterio, "Tabela Geral de Eventos", True, EVENTOINICIAL.Text, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If


  Set Interface = Nothing

End Sub


Public Sub TABELAFILME_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraTabelaFilme(TABELAFILME.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABELAFILME").Value = vHandle
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

  ' Atribuir ESTRUTURAINICIAL E FINAL
  Dim SQLTGE, SQLMASC As Object
  Dim Estrutura As String

  ' Atribuir ESTRUTURAINICIAL
  Set SQLTGE = NewQuery
  SQLTGE.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")
  SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
  SQLTGE.Active = True
  CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = SQLTGE.FieldByName("ESTRUTURA").Value

  ' Atribuir ESTRUTURAFINAL
  SQLTGE.Active = False
  SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
  SQLTGE.Active = True
  Estrutura = SQLTGE.FieldByName("ESTRUTURA").Value
  SQLTGE.Active = False
  Set SQLTGE = Nothing

  ' Completar ESTRUTURAFinal com 99999
  Set SQLMASC = NewQuery
  SQLMASC.Add("SELECT M.MASCARA MASCARA FROM Z_TABELAS T, Z_MASCARAS M")
  SQLMASC.Add("WHERE T.NOME = 'SAM_TGE' AND M.TABELA = T.HANDLE")
  SQLMASC.Active = True
  Estrutura = Estrutura + Mid(SQLMASC.FieldByName("MASCARA").AsString, Len(Estrutura) + 1)
  CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Estrutura
  SQLMASC.Active = False
  Set SQLMASC = Nothing

  If CurrentQuery.FieldByName("TABCOPIAPARA").AsInteger = 1 Then
    If CurrentQuery.FieldByName("PERCENTUALPGTOUS").IsNull Then
      bsShowMessage("Campo '% pagamento US ' obrigatório !", "E")
      CanContinue = False
      Exit Sub
    Else
      If CurrentQuery.FieldByName("PERCENTUALPGTOFILME").IsNull Then
        bsShowMessage("Campo '% pagamento Filme ' obrigatório !", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  End If


  If CurrentQuery.FieldByName("TABCOPIAPARA").AsInteger = 2 Then
    If CurrentQuery.FieldByName("ESTADO").IsNull Then
      bsShowMessage("Campo 'Estado' obrigatório !", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("TABCOPIAPARA").AsInteger = 3 Then
    If CurrentQuery.FieldByName("ESTADO").IsNull Then
      bsShowMessage("Campo 'Estado' obrigatório !", "E")
      CanContinue = False
      Exit Sub
    Else
      If CurrentQuery.FieldByName("MUNICIPIO").IsNull Then
        bsShowMessage("Campo 'Município' obrigatório !", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  End If

  If CurrentQuery.FieldByName("TABCOPIAPARA").AsInteger = 4 Then
    If CurrentQuery.FieldByName("REDERESTRITA").IsNull Then
      bsShowMessage("Campo 'Rede restrita' obrigatório !", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("TABCOPIAPARA").AsInteger = 5 Then
    If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
      bsShowMessage("Campo 'Prestador' obrigatório !", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("TABCOPIAPARA").AsInteger = 6 Then
    If CurrentQuery.FieldByName("REDERESTRITA").IsNull Then
      bsShowMessage("Campo 'Rede restrita' obrigatório !", "E")
      CanContinue = False
      Exit Sub
    Else
      If CurrentQuery.FieldByName("REDERESTRITAPRESTADOR").IsNull Then
        bsShowMessage("Campo 'Prestador' obrigatório !", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  End If

  If CurrentQuery.FieldByName("TABCOPIAPARA").AsInteger = 7 Then
    If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
      bsShowMessage("Campo 'Prestador' obrigatório !", "E")
      CanContinue = False
      Exit Sub
    Else
      If CurrentQuery.FieldByName("MEMBROCORPOCLINICO").IsNull Then
        bsShowMessage("Campo 'Membro corpo-clínico' obrigatório !", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  End If


End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
   If CommandID = "BOTAOPROCESSAR" Then
	BOTAOPROCESSAR_OnClick
   End If
End Sub

Public Sub TABUS_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  Dim vCampos As String
  Dim Interface As Object
  Dim SQL As Object
  Dim vColunas As String
  Dim vCriterio As String
  Dim vDAta As String



  Set Interface = CreateBennerObject("Procura.Procurar")
  ShowPopup = False

  Set SQL = NewQuery
  SQL.Add("SELECT FILTRARTABELAUS FROM SAM_PARAMETROSPRESTADOR")
  SQL.Active = True

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "DESCRICAO|DATAINICIAL|DATAFINAL"


  vDAta = SQLDate(ServerDate)

  vCriterio = ""
  If SQL.FieldByName("FILTRARTABELAUS").AsString = "S" Then

      vCriterio = vCriterio + "SAM_TABUS_VLR.HANDLE IN (SELECT DISTINCT(V.HANDLE) FROM SAM_TABUS_VLR V                           "
      vCriterio = vCriterio + "                          WHERE SAM_TABUS.HANDLE = V.TABELAUS                                     "
      vCriterio = vCriterio + "                            AND V.DATAINICIAL <= " + vDAta + "                                        "
      vCriterio = vCriterio + "                            AND (V.DATAFINAL IS NULL OR V.DATAFINAL >= " + vDAta + ")                 "
      vCriterio = vCriterio + "                        )                                                                         "
  End If

  vColunas = "DESCRICAO|DATAINICIAL|DATAFINAL"
  vCampos = "Descrição da Tabela|Data inicial|Data final"

  vHandle = Interface.Exec(CurrentSystem, "SAM_TABUS|SAM_TABUS_VLR[SAM_TABUS_VLR.TABELAUS = SAM_TABUS.HANDLE]", vColunas, 1, vCampos, vCriterio, "Tabela de US", True, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABUS").Value = vHandle
  End If

  Set Interface = Nothing


End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  Dim Interface As Object
  Dim Handlexx As Long
  Dim vCondicao As String

  CurrentQuery.FieldByName("MEMBROCORPOCLINICO").Value = Null

  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")
  Handlexx = -1
  vCondicao = ""

  Handlexx = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", "PRESTADOR|Z_NOME", 1, "Prestador|Nome", vCondicao, "Tabela de prestadores", True, "")

  If Handlexx > 0 Then
    CurrentQuery.FieldByName("PRESTADOR").Value = Handlexx
  End If


  Set Interface = Nothing

End Sub


Public Sub MEMBROCORPOCLINICO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim Handlexx As Long
  Dim vCondicao As String


  ShowPopup = False

  Set Interface = CreateBennerObject("Procura.Procurar")
  Handlexx = -1

  If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
    bsShowMessage("Informar prestador !", "I")
    Exit Sub
  End If

  vCondicao = "SAM_PRESTADOR_PRESTADORDAENTID.ENTIDADE = " + CurrentQuery.FieldByName("PRESTADOR").AsString

  Handlexx = Interface.Exec(CurrentSystem, "SAM_PRESTADOR_PRESTADORDAENTID|SAM_PRESTADOR[SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PRESTADORDAENTID.PRESTADOR]", _
             "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.Z_NOME|SAM_PRESTADOR_PRESTADORDAENTID.DATAINICIAL|SAM_PRESTADOR_PRESTADORDAENTID.DATAFINAL", _
             1, "Prestador|Nome|Data inicial|Data final", vCondicao, "Tabela de membros do prestador selecionado", True, "")

  If Handlexx > 0 Then
    CurrentQuery.FieldByName("MEMBROCORPOCLINICO").Value = Handlexx
  End If


  Set Interface = Nothing

End Sub

Public Sub TABLE_AfterInsert()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")
  SQL.Active = True

  If SQL.FieldByName("TOTAL").AsInteger = 1 Then
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")
    SQL.Active = True
    CurrentQuery.FieldByName("CONVENIO").Value = SQL.FieldByName("HANDLE").Value
  End If
  Set SQL = Nothing

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

