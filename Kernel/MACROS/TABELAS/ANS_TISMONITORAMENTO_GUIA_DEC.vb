'HASH: 3FAD6BAF0F962DC10AB2A0F7C6881B7D

'#Uses "*bsShowMessage"
Option Explicit

Dim vCodigoCampoComErro As String
Dim vFoiALteradoAlgumCampo As Boolean

Public Sub TABLE_AfterInsert()
  If SessionVar("CODIGOCAMPOCOMERRO") <> "" Then
    CurrentQuery.FieldByName("ROTINAMONITORAMENTOGUIA").AsInteger = CLng(RecordHandleOfTable("ANS_TISMONITORAMENTO_GUIA"))
    If SessionVar("CODIGOCAMPOCOMERRO") = "062" Then
      CurrentQuery.FieldByName("TIPODECLARACAO").AsInteger = 1
    Else
      CurrentQuery.FieldByName("TIPODECLARACAO").AsInteger = 2
    End If
  End If
End Sub

Public Sub TABLE_AfterPost()
  If SessionVar("CODIGOCAMPOCOMERRO") <> "" Then
    Dim component As CSBusinessComponent
    Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Guia.Correcao, Benner.Saude.ANS.Processos")
    component.AddParameter(pdtInteger, CLng(RecordHandleOfTable("ANS_TISMONITORAMENTO_GUIA")))
    component.AddParameter(pdtString, SessionVar("CODIGOCAMPOCOMERRO"))
    component.Execute("RemoverErro")
    Set component = Nothing

    RefreshNodesWithTable("ANS_TISMONITORAMENTO_GUIA")
  End If
End Sub

Public Sub TABLE_AfterScroll()
  BloquearTodosCampos

  If SessionVar("CODIGOCAMPOCOMERRO") <> "" Then
    TIPODECLARACAO.ReadOnly = False
    DECLARACAO.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim qSql As BPesquisa
    Set qSql = NewQuery
    qSql.Add("SELECT E.IDENTIFICADORCAMPO                                          ")
    qSql.Add("  FROM ANS_TISMONITORAMENTO_ERROGUIA  E                              ")
    qSql.Add("  JOIN ANS_TISMONITORAMENTO_GUIA G ON G.HANDLE = E.GUIA              ")
    qSql.Add(" WHERE G.HANDLE = :HANDLE                                            ")
    qSql.Add("   AND E.IDENTIFICADORCAMPO IN ('062', '063')                        ")

    qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAMONITORAMENTOGUIA").AsInteger
    qSql.Active = True

	If Not qSql.EOF Then
      LiberarCamposParaAjuste
    Else
      bsShowMessage("Não é possível alterar as declarações, pois não existe erro em nenhuma delas.", "E")
      CanContinue = False
    End If

    qSql.Active = False
    Set qSql = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If (SessionVar("CODIGOCAMPOCOMERRO") = "063") And (CurrentQuery.FieldByName("TIPODECLARACAO").AsInteger <> 2) Then
    CanContinue = False
    bsShowMessage("O tipo de declaração deve ser de óbito", "E")
  ElseIf (SessionVar("CODIGOCAMPOCOMERRO") = "062") And (CurrentQuery.FieldByName("TIPODECLARACAO").AsInteger <> 1) Then
    CanContinue = False
    bsShowMessage("O tipo de declaração deve ser Nascido Vivo", "E")
  Else
	VerificarSeCampoFoiALterado

	If vFoiALteradoAlgumCampo Then
		Dim qSql As BPesquisa
		Set qSql = NewQuery
		qSql.Add("SELECT HANDLE                                 ")
		qSql.Add("  FROM ANS_TISMONITORAMENTO_GUIA_DEC          ")
		qSql.Add(" WHERE ROTINAMONITORAMENTOGUIA = :GUIA ")
		qSql.Add("   AND HANDLE <> :HANDLE                      ")

		qSql.ParamByName("GUIA").AsInteger = CurrentQuery.FieldByName("ROTINAMONITORAMENTOGUIA").AsInteger
		qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		qSql.Active = True

		If Not qSql.EOF Then
		  If bsShowMessage("Existe outras declarações vinculadas a este procedimento. Deseja alterar as demais declarações?", "Q") = vbYes Then
		    Exit Sub
		  End If
		End If

		ExcluirErroDoProcedimento("062")
		ExcluirErroDoProcedimento("063")

		qSql.Active = False
		Set qSql = Nothing

		RefreshNodesWithTable("ANS_TISMONITORAMENTO_GUIA_DEC")
	ElseIf CurrentQuery.State = 3  Then
		bsShowMessage("Não foi realizado alteração nos campos do pacote","E")
		CanContinue = False
	End If
  End If
End Sub

Public Sub BloquearTodosCampos
  TIPODECLARACAO.ReadOnly = True
  DECLARACAO.ReadOnly = True
End Sub

Public Sub LiberarCamposParaAjuste
  TIPODECLARACAO.ReadOnly = False
  DECLARACAO.ReadOnly = False
End Sub

Public Function VerificarSeCampoFoiALterado()

  Dim registroOriginal As BPesquisa
  Set registroOriginal = NewQuery

  registroOriginal.Add("SELECT *                              ")
  registroOriginal.Add("  FROM ANS_TISMONITORAMENTO_GUIA_DEC ")
  registroOriginal.Add(" WHERE HANDLE = :HANDLE               ")

  registroOriginal.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("DECLARACAOORIGEM").AsInteger
  registroOriginal.Active = True

  vFoiALteradoAlgumCampo = False

  If Not registroOriginal.EOF Then
    If registroOriginal.FieldByName("DECLARACAO").AsString <> CurrentQuery.FieldByName("DECLARACAO").AsString _
    Or registroOriginal.FieldByName("TIPODECLARACAO").AsInteger <> CurrentQuery.FieldByName("TIPODECLARACAO").AsInteger Then
	  vFoiALteradoAlgumCampo = True
	End If
  End If

  registroOriginal.Active = False
  Set registroOriginal = Nothing

End Function

Public Function ExcluirErroDoProcedimento(campoASerExcluido As String)
  Dim component As CSBusinessComponent
  Set component = BusinessComponent.CreateInstance("Benner.Saude.ANS.Processos.Monitoramento.Guia.Correcao, Benner.Saude.ANS.Processos")
  component.AddParameter(pdtInteger, CurrentQuery.FieldByName("ROTINAMONITORAMENTOGUIA").AsInteger)
  component.AddParameter(pdtString, campoASerExcluido)
  component.Execute("RemoverErro")
  Set component = Nothing
End Function
