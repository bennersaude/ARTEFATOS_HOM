'HASH: D6B89E89884DF4D7142E59B133ACDD01

'#Uses "*bsShowMessage"
Option Explicit
Dim vCodigoCampoComErro As String
Dim vFoiALteradoCodigoTabela As Boolean
Dim vFoiALteradoQuantidade As Boolean

Public Sub TABLE_AfterScroll()
  BloquearTodosCampos

  If WebMode Then
    CODIGOTABELA.WebLocalWhere = "A.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
  Else
    CODIGOTABELA.LocalWhere = "VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If CurrentQuery.FieldByName("PACOTEORIGEM").AsInteger > 0 Then
    Dim qSql As BPesquisa
    Set qSql = NewQuery
    qSql.Add("SELECT E.IDENTIFICADORCAMPO                                          ")
    qSql.Add("  FROM ANS_TISMONITORAMENTO_ERROPROC  E                              ")
    qSql.Add("  JOIN ANS_TISMONITORAMENTO_GUIA_PROC P ON P.HANDLE = E.PROCEDIMENTO ")
    qSql.Add(" WHERE P.HANDLE = :HANDLE                                            ")
    qSql.Add("   AND E.IDENTIFICADORCAMPO IN ('077', '078', '079')                 ")

    qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PROCMONITORAMENTO").AsInteger
    qSql.Active = True

    If Not qSql.EOF Then
      LiberarCamposParaAjuste
    Else
      bsShowMessage("Não é possível alterar o pacote, pois não existe erro em nenhum deles", "E")
      CanContinue = False
    End If

    qSql.Active = False
    Set qSql = Nothing
  Else
    CanContinue = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  VerificarSeCampoFoiALterado

  If vFoiALteradoQuantidade Or vFoiALteradoCodigoTabela Then
    Dim qSql As BPesquisa
    Set qSql = NewQuery
    qSql.Add("SELECT HANDLE                                 ")
    qSql.Add("  FROM ANS_TISMONITORAMENTO_PACOTE            ")
    qSql.Add(" WHERE PROCMONITORAMENTO = :PROCMONITORAMENTO ")
    qSql.Add("   AND HANDLE <> :HANDLE                      ")

    qSql.ParamByName("PROCMONITORAMENTO").AsInteger = CurrentQuery.FieldByName("PROCMONITORAMENTO").AsInteger
    qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSql.Active = True

    If Not qSql.EOF Then
      If bsShowMessage("Existe outros pacotes vinculados a este procedimento. Deseja alterar os demais itens do pacote?", "Q") = vbYes Then
        Exit Sub
      End If
    End If

    If vFoiALteradoCodigoTabela Then
      ExcluirErroDoProcedimento("078")
      ExcluirErroDoProcedimento("077")
    End If

    If vFoiALteradoQuantidade Then
      ExcluirErroDoProcedimento("079")
    End If

    qSql.Active = False
    Set qSql = Nothing

    RefreshNodesWithTable("ANS_TISMONITORAMENTO_GUIA_PROC")
  Else
    bsShowMessage("Não foi realizado alteração nos campos do pacote","E")
    CanContinue = False
  End If

End Sub

Public Sub BloquearTodosCampos
  CODIGOTABELA.ReadOnly = True
  PROCEDIMENTO.ReadOnly = True
  QUANTIDADE.ReadOnly = True
  ENVIOCONSOLIDADO.ReadOnly = True
End Sub

Public Sub LiberarCamposParaAjuste
  CODIGOTABELA.ReadOnly = False
  QUANTIDADE.ReadOnly = False
  ENVIOCONSOLIDADO.ReadOnly = False
End Sub

Public Function ExcluirErroDoProcedimento(campoASerExcluido As String)
  Dim component As CSBusinessComponent
  Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Procedimento.Correcao, Benner.Saude.ANS.Processos")
  component.AddParameter(pdtInteger, CurrentQuery.FieldByName("PROCMONITORAMENTO").AsInteger)
  component.AddParameter(pdtString, campoASerExcluido)
  component.Execute("RemoverErro")
  Set component = Nothing
End Function

Public Function VerificarSeCampoFoiALterado()

  Dim registroOriginal As BPesquisa
  Set registroOriginal = NewQuery

  registroOriginal.Add("SELECT *                              ")
  registroOriginal.Add("  FROM ANS_TISMONITORAMENTO_PACOTE ")
  registroOriginal.Add(" WHERE HANDLE = :HANDLE               ")

  registroOriginal.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PACOTEORIGEM").AsInteger
  registroOriginal.Active = True

  vFoiALteradoCodigoTabela = False
  vFoiALteradoQuantidade = False

  If Not registroOriginal.EOF Then
    If registroOriginal.FieldByName("CODIGOTABELA").AsInteger <> CurrentQuery.FieldByName("CODIGOTABELA").AsInteger Then
	  vFoiALteradoCodigoTabela = True
	End If
    If registroOriginal.FieldByName("QUANTIDADE").AsFloat <> CurrentQuery.FieldByName("QUANTIDADE").AsFloat Then
	  vFoiALteradoQuantidade = True
	End If
    If registroOriginal.FieldByName("ENVIOCONSOLIDADO").AsString <> CurrentQuery.FieldByName("ENVIOCONSOLIDADO").AsString Then
      vFoiALteradoCodigoTabela = True
    End If
  End If

  registroOriginal.Active = False
  Set registroOriginal = Nothing

End Function
