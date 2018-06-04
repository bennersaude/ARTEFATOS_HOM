'HASH: 89834D97D9DFE05BA7422838782D1973
'Macro: SAM_TIPOAUTORIZ

'#Uses "*bsShowMessage"
Option Explicit

Public Sub CONDICAOATENDIMENTO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_CONDATENDIMENTO")
End Sub

Public Sub FINALIDADEATENDIMENTO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_FINALIDADEATENDIMENTO")
End Sub

Public Sub LOCALATENDIMENTO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_LOCALATENDIMENTO")
End Sub

Public Sub OBJETIVOTRATAMENTO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_OBJTRATAMENTO")
End Sub

Public Sub REGIMEATENDIMENTO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_REGIMEATENDIMENTO")
End Sub

Public Sub RELATORIO_OnBtnClick()
  If CurrentQuery.State = 1 Then
    bsShowMessage("Registro deve estar em edição ou inserção", "I")
  Else
    Dim vHandleRelatorio As Long
    Dim Interface As Object
    Dim vCampos As String
    Dim vColunas As String
    Dim vCriterio As String
    Dim vTabela As String
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CODIGO, NOME FROM R_RELATORIOS WHERE HANDLE = :H")
    Set Interface = CreateBennerObject("Procura.Procurar")
    vColunas = "CODIGO|NOME"
    vCriterio = ""
    vCampos = "Código |Relatório"
    vTabela = "R_RELATORIOS"
    'vHandleRelatorio =Interface.Exec(vTabela,vColunas,2,vCampos," CODIGO LIKE 'AUT%'","Relatório",True,"")
    vHandleRelatorio = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, " ((NOME LIKE 'AUT%') OR (NOME LIKE 'TISS%')) ", "Relatório", True, "")
    If vHandleRelatorio >0 Then
      SQL.Active = False
      SQL.ParamByName("H").Value = vHandleRelatorio
      SQL.Active = True
      CurrentQuery.FieldByName("RELATORIO").Value = SQL.FieldByName("CODIGO").AsString
      NOMERELATORIO.Text = SQL.FieldByName("NOME").AsString
    End If
    Set Interface = Nothing
    Set SQL = Nothing
  End If
End Sub

Public Sub TABLE_AfterScroll()

  TISSTIPOSOLICITACAO_OnChange

  If WebMode Then
    RELATORIOAUTORIZACAO.WebLocalWhere =  "A.HANDLE IN (SELECT RELATORIO FROM SAM_RELATORIOAUT)"
  Else
    RELATORIOAUTORIZACAO.LocalWhere =  "HANDLE IN (SELECT RELATORIO FROM SAM_RELATORIOAUT)"
  End If

  Dim vAux As String
  vAux = "(SELECT FINALIDADEATENDIMENTO FROM SAM_TIPOAUTORIZ_FINATEND WHERE TIPOAUTORIZ = " + Str(CurrentQuery.FieldByName("HANDLE").AsInteger) + ")"
  FINALIDADEATENDIMENTO.LocalWhere = "NOT EXISTS " + vAux + " OR HANDLE IN " + vAux
  vAux = "(SELECT CONDICAOATENDIMENTO FROM SAM_TIPOAUTORIZ_CONDICAOATEND WHERE TIPOAUTORIZ = " + Str(CurrentQuery.FieldByName("HANDLE").AsInteger) + ")"
  CONDICAOATENDIMENTO.LocalWhere = "NOT EXISTS " + vAux + " OR HANDLE IN " + vAux
  vAux = "(SELECT LOCALATENDIMENTO FROM SAM_TIPOAUTORIZ_LOCALATEND WHERE TIPOAUTORIZ = " + Str(CurrentQuery.FieldByName("HANDLE").AsInteger) + ")"
  LOCALATENDIMENTO.LocalWhere = "NOT EXISTS " + vAux + " OR HANDLE IN " + vAux
  vAux = "(SELECT REGIMEATENDIMENTO FROM SAM_TIPOAUTORIZ_REGIMEATEND WHERE TIPOAUTORIZ = " + Str(CurrentQuery.FieldByName("HANDLE").AsInteger) + ")"
  REGIMEATENDIMENTO.LocalWhere = "NOT EXISTS " + vAux + " OR HANDLE IN " + vAux
  vAux = "(SELECT OBJETIVOTRATAMENTO FROM SAM_TIPOAUTORIZ_OBJTRATAMENTO WHERE TIPOAUTORIZ = " + Str(CurrentQuery.FieldByName("HANDLE").AsInteger) + ")"
  OBJETIVOTRATAMENTO.LocalWhere = "NOT EXISTS " + vAux + " OR HANDLE IN " + vAux
  vAux = "(SELECT TIPOTRATAMENTO FROM SAM_TIPOAUTORIZ_TIPOTRATAMENTO WHERE TIPOAUTORIZ = " + Str(CurrentQuery.FieldByName("HANDLE").AsInteger) + ")"
  TIPOTRATAMENTO.LocalWhere = "NOT EXISTS " + vAux + " OR HANDLE IN " + vAux
End Sub

Public Sub TABCOBRANCAPF_OnChange()
  CurrentQuery.UpdateRecord
  If CurrentQuery.FieldByName("TABCOBRANCAPF").AsInteger <> 1 Then
	  Dim SQL As Object
	  Set SQL = NewQuery
	  SQL.Add("SELECT * FROM SAM_TIPOAUTORIZ_FRQINTERNACAO             ")
	  SQL.Add("	WHERE TIPOAUTORIZ = :TIPOAUTORIZ                       ")
	  SQL.Add("	And ((COMPETENCIAINICIAL IS NOT NULL AND COMPETENCIAFINAL = NULL) ")
	  SQL.Add("	OR ( :DATA BETWEEN COMPETENCIAINICIAL AND COMPETENCIAFINAL)) ")
	  SQL.ParamByName("DATA").AsDateTime = Now
	  SQL.ParamByName("TIPOAUTORIZ").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	  SQL.Active = True
	  SQL.First
	  If Not SQL.EOF Then
	    CurrentQuery.Cancel
		bsShowMessage("Tipo de autorização com franquia de internação configurada não é permitido ter “Cobrança de PF” diferente de “No proc. contas”.", "E")
	  End If
	  Set SQL = Nothing
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("TISSTIPOSOLICITACAO").AsString <> "3" Then
    CurrentQuery.FieldByName("GERARPLANOTRATAMENTO").AsString = "N"
  End If
  ' SMS - 50420
  Dim Interface As Object
  Dim Linha As String
  Dim SQL As Object

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_TIPOAUTORIZ", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CODIGO", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
  Set Interface = Nothing
  ' Fim SMS - 50420

  'SMS 81867 - Débora Rebello - 21/05/2007 - inicio
  If (CurrentQuery.FieldByName("DATAFINAL").AsString <> "") Then
    Linha = VerificarSePodeEncerrar

    If Linha <> "" Then
      CanContinue = False
      bsShowMessage(Linha, "E")
      Exit Sub
    Else
      CanContinue = True
    End If
  End If
  'SMS 81867 - Débora Rebello - 21/05/2007 - fim


  If CurrentQuery.FieldByName("EXECUTORRESUMO").AsString = "N" And _
                               CurrentQuery.FieldByName("RECEBEDORRESUMO").AsString = "N" And _
                               CurrentQuery.FieldByName("LOCALEXECUCAORESUMO").AsString = "N" And _
                               CurrentQuery.FieldByName("SOLICITANTERESUMO").AsString = "N" Then
    CanContinue = False
    bsShowMessage("É obrigatório marcar pelo menos um tipo de prestador para ser exibido no Resumo da Autorização", "E")
  End If

  ' Eduardo -30/07/2004 -SMS 30513
  ' Código temporário até a aplicação do runner 5.3.63
  If FINALIDADEATENDIMENTO.Text = "???" Then
    CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").Clear
  End If
  If CONDICAOATENDIMENTO.Text = "???" Then
    CurrentQuery.FieldByName("CONDICAOATENDIMENTO").Clear
  End If
  If LOCALATENDIMENTO.Text = "???" Then
    CurrentQuery.FieldByName("LOCALATENDIMENTO").Clear
  End If
  If REGIMEATENDIMENTO.Text = "???" Then
    CurrentQuery.FieldByName("REGIMEATENDIMENTO").Clear
  End If
  If OBJETIVOTRATAMENTO.Text = "???" Then
    CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").Clear
  End If
  If TIPOTRATAMENTO.Text = "???" Then
    CurrentQuery.FieldByName("TIPOTRATAMENTO").Clear
  End If
  ' fim SMS 30513

  '--- Autorizador externo Durval --- SMS 41618
  'Se marcou como padrão, deve escolher um tipo
  If CurrentQuery.FieldByName("PADRAOAUTORIZADOREXTERNO").AsString = "S" And _
                              CurrentQuery.FieldByName("TIPOAUTORIZACAOEXTERNA").AsString = "" Then
    CanContinue = False
    bsShowMessage("Nenhum tipo de autorização externa escolhida para ser a padrão.", "E")
  End If
  'Se não marcou como padrão e é a única incidência do tipo de autorização, marca como padrão
  If CurrentQuery.FieldByName("PADRAOAUTORIZADOREXTERNO").AsString = "N" And _
                              CurrentQuery.FieldByName("TIPOAUTORIZACAOEXTERNA").AsString <> "" Then

    Set SQL = NewQuery
    SQL.Clear
    SQL.Add("SELECT COUNT(*) T FROM SAM_TIPOAUTORIZ WHERE HANDLE <> :HANDLE AND TIPOAUTORIZACAOEXTERNA = :TIPOAUTORIZACAOEXTERNA AND PADRAOAUTORIZADOREXTERNO = 'S'")
    SQL.ParamByName("TIPOAUTORIZACAOEXTERNA").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAOEXTERNA").Value
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.Active = True
    If SQL.FieldByName("T").AsInteger = 0 Then
      SQL.Active = False
      bsShowMessage("Autorização externa marcada como padrão.", "I")
      CurrentQuery.FieldByName("PADRAOAUTORIZADOREXTERNO").AsString = "S"
    End If
    Set SQL = Nothing
  End If
  '--- Fim Autorizador externo ---

End Sub

Public Sub TIPOTRATAMENTO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_TIPOTRATAMENTO")
End Sub


Public Sub INSERIR(pTabela, pCampo As String)

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT COUNT(*) T FROM " + pTabela + " WHERE TIPOAUTORIZ = :TIPOAUTORIZ And " + pCampo + " = :" + pCampo)
  SQL.ParamByName("TIPOAUTORIZ").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName(pCampo).Value = CurrentQuery.FieldByName(pCampo).AsInteger
  SQL.Active = True
  If SQL.FieldByName("T").AsInteger = 0 Then
    SQL.Active = False
    SQL.Clear
    SQL.Add("INSERT INTO " + pTabela + "(HANDLE, TIPOAUTORIZ," + pCampo + ") VALUES")
    SQL.Add("(:HANDLE,:TIPOAUTORIZ,:" + pCampo + ")")
    SQL.ParamByName("HANDLE").Value = NewHandle(pTabela)
    SQL.ParamByName("TIPOAUTORIZ").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ParamByName(pCampo).Value = CurrentQuery.FieldByName(pCampo).AsInteger
    SQL.ExecSQL
  End If
  Set SQL = Nothing
End Sub


Public Sub TABLE_AfterPost()

  RefreshNodesWithTable("SAM_TIPOAUTORIZ")

  '--- Autorizador externo Durval --- SMS 41618
  Dim SQL As Object
  Set SQL = NewQuery

  'Se marcou como padrão e escolheu um tipo, garante que este é o único deste tipo que é padrão.
  If CurrentQuery.FieldByName("PADRAOAUTORIZADOREXTERNO").AsString = "S" And _
                              CurrentQuery.FieldByName("TIPOAUTORIZACAOEXTERNA").AsString <> "" Then
    SQL.Clear
    SQL.Add("UPDATE SAM_TIPOAUTORIZ SET PADRAOAUTORIZADOREXTERNO = 'N' ")
    SQL.Add(" WHERE HANDLE <> :HANDLE AND TIPOAUTORIZACAOEXTERNA = :TIPOAUTORIZACAOEXTERNA AND TISSTIPOSOLICITACAO = :TISSTIPOSOLICITACAO")
    SQL.ParamByName("TIPOAUTORIZACAOEXTERNA").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAOEXTERNA").Value
    SQL.ParamByName("TISSTIPOSOLICITACAO").Value = CurrentQuery.FieldByName("TISSTIPOSOLICITACAO").Value
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.ExecSQL
    Set SQL = Nothing
  End If
  'Se não marcou como padrão e é a única incidência do tipo de autorização, marca como padrão
  If CurrentQuery.FieldByName("PADRAOAUTORIZADOREXTERNO").AsString = "N" And _
                              CurrentQuery.FieldByName("TIPOAUTORIZACAOEXTERNA").AsString <> "" Then
    SQL.Clear
    SQL.Add("SELECT COUNT(*) T FROM SAM_TIPOAUTORIZ WHERE TIPOAUTORIZACAOEXTERNA = :TIPOAUTORIZACAOEXTERNA AND HANDLE <> :HANDLE")
    SQL.ParamByName("TIPOAUTORIZACAOEXTERNA").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAOEXTERNA").Value
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.Active = True
    If SQL.FieldByName("T").AsInteger = 0 Then
      SQL.Active = False
      SQL.Clear
      SQL.Add("UPDATE SAM_TIPOAUTORIZ SET PADRAOAUTORIZADOREXTERNO = 'S' WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
      SQL.ExecSQL
    End If
    Set SQL = Nothing
  End If
  '--- Fim Autorizador externo ---
End Sub

Function VerificarSePodeEncerrar() As String
  'SMS 81867 - Débora Rebello - 21/05/2007
  Dim SQL As Object
  Set SQL = NewQuery

  VerificarSePodeEncerrar = ""

  'Verificando se o tipo de autorização a ser encerrado está como tipo de autorização padrão de algum usuário
  SQL.Clear
  SQL.Add("SELECT TIPOAUTORIZACAOPADRAO                     ")
  SQL.Add("  FROM Z_GRUPOUSUARIOS U                         ")
  SQL.Add(" WHERE TIPOAUTORIZACAOPADRAO = :TIPOAUTORIZ      ")
  SQL.Add("UNION                                            ")
  'Verificando se o tipo de autorização a ser encerrado está como tipo de autorização padrão de algum grupo de segurança
  SQL.Add("SELECT TIPOAUTORIZACAOPADRAO                     ")
  SQL.Add("  FROM Z_GRUPOS G                                ")
  SQL.Add(" WHERE TIPOAUTORIZACAOPADRAO = :TIPOAUTORIZ      ")
  SQL.Add("UNION                                            ")
  'Verificando se o tipo de autorização a ser encerrado está como tipo de autorização padrão dos parâmetros gerais de atendimento
  SQL.Add("SELECT AUTOTIPOAUTORIZPADRAO                     ")
  SQL.Add("  FROM SAM_PARAMETROSATENDIMENTO                 ")
  SQL.Add(" WHERE AUTOTIPOAUTORIZPADRAO = :TIPOAUTORIZ      ")
  SQL.ParamByName("TIPOAUTORIZ").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If (Not SQL.EOF) Then
    VerificarSePodeEncerrar = "Não é possível encerrar a vigência desse tipo de autorização porque ele está configurado " + _
       "como tipo de autorização padrão de algum usuário (Adm/Usuários) ou de algum grupo de segurança (Adm/Grupos de segurança) "  + _
       "ou dos parâmetros gerais de atendimento(Adm/Parâmetros gerais/Atendimento). É necessário retirar todas essas configurações "  + _
       "antes de encerrar a vigência."
  End If

  Set SQL = Nothing

End Function

Public Sub TISSTIPOSOLICITACAO_OnChange()

  If CurrentQuery.State = 2 Then
    CurrentQuery.UpdateRecord
  End If
  If CurrentQuery.FieldByName("TISSTIPOSOLICITACAO").AsString = "3" Then
    GERARPLANOTRATAMENTO.Visible = True
  Else
    GERARPLANOTRATAMENTO.Visible = False
  End If
End Sub
