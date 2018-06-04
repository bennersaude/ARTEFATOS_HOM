'HASH: 4D3965440CB7B1463F74A85FD580B60E
'MACRO='Macro: SAM_GUIA_EVENTOS_NEGACAO

'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"
'#Uses "*VerificaPermissaoEdicaoTriagem"
'#Uses "*RecordHandleOfTableInterfacePEG"
'#Uses "*RefreshNodesWithTableInterfacePEG"

Option Explicit

Dim CANCONTINUE As Boolean
Dim alterouRevisada As Boolean
Dim vSituacaoAnteriorGuia As String
Dim vSituacaoAnteriorEvento As String

Public Function VERIFICAINCOMP As Boolean
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO FROM SAM_PEG WHERE HANDLE=:PEG")
  q1.ParamByName("PEG").Value = RecordHandleOfTableInterfacePEG("SAM_PEG")
  q1.Active = True
  VERIFICAINCOMP = True
  If q1.FieldByName("SITUACAO").AsString = "3" Then 'PRONTO
    Dim interface As Object
    Set interface = CreateBennerObject("SAMINCOMP.CHECK")
    interface.INICIALIZAR(CurrentSystem)
    VERIFICAINCOMP = interface.EXCLUIRINCOMPATIBILIDADES(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger)
    interface.FINALIZAR
    Set interface = Nothing
  End If
  Set q1 = Nothing
End Function

Public Sub VERIFICAJAPAGO(CANCONTINUE As Boolean)
  Dim SITUACAO As String
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Clear
  q1.Add("SELECT G.SITUACAO FROM SAM_GUIA G WHERE G.HANDLE=:GUIA")
  q1.ParamByName("GUIA").Value = RecordHandleOfTableInterfacePEG("SAM_GUIA")
  q1.Active = True
  SITUACAO = q1.FieldByName("SITUACAO").AsString
  q1.Active = False
  Set q1 = Nothing
  If SITUACAO = "4" Then
    CANCONTINUE = False
  Else
    CANCONTINUE = True
  End If
End Sub

Public Sub BOTAORECONSIDERACAO_OnClick()

  Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RetornaFilialPEG, "A")
  If aux = False Then
    Exit Sub
  End If

  'Necessário verificar se o usuário pode reconsiderar a negação, ou seja, se esta negação está cadastrada na carga de negações que fica embaixo do usuário
  'Inicio
  Dim qUsuarioNegacao As Object
  Set qUsuarioNegacao = NewQuery
  Dim vQtdeUsuarioNegacao As Integer

  qUsuarioNegacao.Clear
  qUsuarioNegacao.Add("SELECT COUNT(1) QTDE")
  qUsuarioNegacao.Add("  FROM SAM_USUARIO_MOTIVONEGACAO")
  qUsuarioNegacao.Add(" WHERE USUARIO = :HUSUARIO")
  qUsuarioNegacao.Add("   AND MOTIVONEGACAO = :HMOTIVONEGACAO")
  qUsuarioNegacao.ParamByName("HUSUARIO").AsInteger = CurrentUser
  qUsuarioNegacao.ParamByName("HMOTIVONEGACAO").AsInteger = CurrentQuery.FieldByName("MOTIVONEGACAO").AsInteger
  qUsuarioNegacao.Active = True

  vQtdeUsuarioNegacao = qUsuarioNegacao.FieldByName("QTDE").AsInteger
  Set qUsuarioNegacao = Nothing

  If vQtdeUsuarioNegacao = 0 Then 'Usuário não pode reconsiderar negação, pois, nào está cadastrada na tabela de negações abaixo do usuário. (Não no grupo de segurança).

    'Aqui será verificado se esta negação nào está cadastrada no grupo de segurança do usuário corrente
    'Início
    Dim qGrupoSeguranca As Object
    Set qGrupoSeguranca = NewQuery
    Dim vQtdeGrupoSeguranca As Integer

    qGrupoSeguranca.Clear
    qGrupoSeguranca.Add("SELECT SUM(X.QTDE) QTDE")
    qGrupoSeguranca.Add("  FROM (")
    qGrupoSeguranca.Add("        SELECT COUNT(1) QTDE ")
    qGrupoSeguranca.Add("          FROM SAM_GRUPO_MOTIVONEGACAO   GM, ")
    qGrupoSeguranca.Add("               Z_GRUPOUSUARIOS          U")
    qGrupoSeguranca.Add("         WHERE GM.GRUPO = U.GRUPO ")
    qGrupoSeguranca.Add("           AND GM.MOTIVONEGACAO = :HMOTIVONEGACAO")
    qGrupoSeguranca.Add("           AND U.HANDLE = :HUSUARIO")
    qGrupoSeguranca.Add("         UNION ALL")
    qGrupoSeguranca.Add("        SELECT COUNT(1) QTDE ")
    qGrupoSeguranca.Add("          FROM SAM_GRUPO_MOTIVONEGACAO   GM, ")
    qGrupoSeguranca.Add("               Z_GRUPOUSUARIOGRUPOS    UG	")
    qGrupoSeguranca.Add("         WHERE UG.GRUPOADICIONADO = GM.GRUPO")
    qGrupoSeguranca.Add("           AND GM.MOTIVONEGACAO = :HMOTIVONEGACAO")
    qGrupoSeguranca.Add("           AND UG.USUARIO = :HUSUARIO")
    qGrupoSeguranca.Add("       ) X")
    qGrupoSeguranca.ParamByName("HMOTIVONEGACAO").AsInteger = CurrentQuery.FieldByName("MOTIVONEGACAO").AsInteger
    qGrupoSeguranca.ParamByName("HUSUARIO").AsInteger = CurrentUser
    qGrupoSeguranca.Active = True

    vQtdeGrupoSeguranca = qGrupoSeguranca.FieldByName("QTDE").AsInteger
    Set qGrupoSeguranca = Nothing

    If vQtdeGrupoSeguranca = 0 Then 'não pode ser reconsiderada por este usuário
      bsShowMessage("Esta negação não pode ser reconsiderada pelo usuário.", "I")
      Exit Sub
    End If
    'FInal

  End If
  'Final

  If VERIFICAINCOMP = False Then
    CANCONTINUE = False
    bsShowMessage("Existe incompatibilidade glosada para este evento", "I")
    Exit Sub
  End If

  VERIFICAJAPAGO(CANCONTINUE)
  If CANCONTINUE = False Then
    bsShowMessage("O Evento já foi pago e não pode ser modificado", "I")
    Exit Sub
  End If

  If Not InTransaction Then StartTransaction

  Dim q1 As Object
  Dim recon As Boolean
  recon = False
  Set q1 = NewQuery

  If CurrentQuery.FieldByName("DATAHORALIBERACAO").IsNull Then
    q1.Clear
    q1.Add("UPDATE SAM_GUIA_EVENTOS_NEGACAO SET NEGACAOREVISADA=:NEGACAOREVISADA, DATAHORALIBERACAO=:DATAHORALIBERACAO, USUARIOLIBERACAO=:USUARIOLIBERACAO,USUARIOREVISOR=:USUARIOLIBERACAO, DATAHORAREVISAO=:DATAHORALIBERACAO WHERE HANDLE=:HANDLE")
    q1.ParamByName("USUARIOLIBERACAO").AsInteger = CurrentUser
    q1.ParamByName("DATAHORALIBERACAO").AsDateTime = ServerNow
    q1.ParamByName("NEGACAOREVISADA").AsString = "S"
    recon = True
  Else
    q1.Clear
    q1.Add("UPDATE SAM_GUIA_EVENTOS_NEGACAO SET DATAHORALIBERACAO=NULL, USUARIOLIBERACAO=NULL, USUARIOREVISOR=NULL, DATAHORAREVISAO=NULL, NEGACAOREVISADA=:NEGACAOREVISADA WHERE HANDLE=:HANDLE")
    q1.ParamByName("NEGACAOREVISADA").Value = "N" 'SMS 68660 - Marcelo Barbosa - 09/10/2006
  End If
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  q1.ExecSQL
  q1.Clear

  q1.Add("UPDATE SAM_GUIA_EVENTOS SET SITUACAO='2' WHERE HANDLE=" + Str(RecordHandleOfTableInterfacePEG("SAM_GUIA_EVENTOS")))
  q1.ExecSQL
  q1.Clear
  q1.Add("UPDATE SAM_GUIA SET SITUACAO='2' WHERE HANDLE=" + Str(RecordHandleOfTableInterfacePEG("SAM_GUIA")))
  q1.ExecSQL
  q1.Clear
  q1.Add("UPDATE SAM_PEG SET SITUACAO='2' WHERE HANDLE=" + Str(RecordHandleOfTableInterfacePEG("SAM_PEG")))
  q1.ExecSQL

  If InTransaction Then Commit

  q1.Clear
  q1.Add("SELECT COUNT(HANDLE) NREC FROM SAM_GUIA_EVENTOS_NEGACAO WHERE GUIAEVENTO=:GUIAEVENTO AND NEGACAOREVISADA='N'")
  q1.ParamByName("GUIAEVENTO").AsInteger = RecordHandleOfTableInterfacePEG("SAM_GUIA_EVENTOS")
  q1.Active = True
  If q1.FieldByName("NREC").AsInteger = 0 Then
  	CriaTabelaTemporariaSqlServer
    Dim interface As Object
    Set interface = CreateBennerObject("SAMPEG.Rotinas")
    interface.RevisarEvento(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, "REVISAR", True)
    Set interface = Nothing
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If

  CurrentQuery.Active = False
  CurrentQuery.Active = True
  ChecarRefresh

  Set q1 = Nothing
End Sub

Public Sub NEGACAOREVISADA_OnChange()
  alterouRevisada = True
End Sub

Public Sub TABLE_AfterCommitted()

	ChecarRefresh

End Sub

Public Sub TABLE_AfterPost()

	Dim interface As Object

  	Set interface = CreateBennerObject("BSPro000.Rotinas")

  	interface.RevisarEvento(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, "REVISAR", True)

  	Set interface = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  Dim hEventoGuia As Long

  If CurrentQuery.FieldByName("GUIAEVENTO").AsInteger <= 0 Then
  	hEventoGuia = RecordHandleOfTableInterfacePEG("SAM_GUIA_EVENTOS")
  Else
  	hEventoGuia = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
  End If

  alterouRevisada = False
  If vSituacaoAnteriorEvento = "" Then
    Dim qSituacaoEvento As Object
    Set qSituacaoEvento = NewQuery
    qSituacaoEvento.Clear
    qSituacaoEvento.Add("SELECT SITUACAO             ")
    qSituacaoEvento.Add("  FROM SAM_GUIA_EVENTOS     ")
    qSituacaoEvento.Add(" WHERE HANDLE = :GUIAEVENTO ")
    qSituacaoEvento.ParamByName("GUIAEVENTO").AsInteger = hEventoGuia
    qSituacaoEvento.Active = True
    vSituacaoAnteriorEvento = qSituacaoEvento.FieldByName("SITUACAO").AsString
    Set qSituacaoEvento = Nothing
  End If
  'SMS 68660 - Marcelo Barbosa - 09/10/2006
  Dim qSituacao As Object
  Set qSituacao = NewQuery
  qSituacao.Clear
  qSituacao.Add("SELECT SITUACAO             ")
  qSituacao.Add("  FROM SAM_GUIA_EVENTOS     ")
  qSituacao.Add(" WHERE HANDLE = :GUIAEVENTO ")
  qSituacao.ParamByName("GUIAEVENTO").AsInteger = hEventoGuia
  qSituacao.Active = True
  If (qSituacao.FieldByName("SITUACAO").AsString = "3") Or (qSituacao.FieldByName("SITUACAO").AsString = "4") Or (Not VerificarPermissaoUsuarioPegTriado(False)) Then
    BOTAORECONSIDERACAO.Enabled = False
  Else
    BOTAORECONSIDERACAO.Enabled = True
  End If
  Set qSituacao = Nothing
  'Fim - SMS 68660
End Sub

Public Sub TABLE_BeforeDelete(CANCONTINUE As Boolean)

  If Not VerificarPermissaoUsuarioPegTriado(False) Then
  	bsShowMessage("Alteração não permitida para este usuário.", "I")
 	CANCONTINUE = False
 	Exit Sub
  End If

  CANCONTINUE = CheckFilialProcessamento(CurrentSystem, RetornaFilialPEG, "E")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CANCONTINUE = False Then
    Exit Sub
  End If

  If VERIFICAINCOMP = False Then
    CANCONTINUE = False
    bsShowMessage("Existe incompatibilidade glosada para este evento", "E")
    Exit Sub
  End If

  VERIFICAJAPAGO(CANCONTINUE)
  If CANCONTINUE = False Then
    bsShowMessage("O Evento já foi pago e não pode ser modificado", "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CANCONTINUE As Boolean)

  If Not VerificarPermissaoUsuarioPegTriado(False) Then
   	bsShowMessage("Alteração não permitida para este usuário.", "I")
 	CANCONTINUE = False
 	Exit Sub
  End If

  CANCONTINUE = CheckFilialProcessamento(CurrentSystem, RetornaFilialPEG, "A")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CANCONTINUE = False Then
    Exit Sub
  End If

  If VERIFICAINCOMP = False Then
    CANCONTINUE = False
    bsShowMessage("Existe incompatibilidade glosada para este evento", "E")
    Exit Sub
  End If

  VERIFICAJAPAGO(CANCONTINUE)
  If CANCONTINUE = False Then
    bsShowMessage("O Evento já foi pago e não pode ser modificado", "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CANCONTINUE As Boolean)

  If Not VerificarPermissaoUsuarioPegTriado(False) Then
   	bsShowMessage("Alteração não permitida para este usuário.", "I")
 	CANCONTINUE = False
 	Exit Sub
  End If

  CANCONTINUE = CheckFilialProcessamento(CurrentSystem, RetornaFilialPEG, "I")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CANCONTINUE = False Then
    Exit Sub
  End If

  If VERIFICAINCOMP = False Then
    bsShowMessage("Existe incompatibilidade glosada para este evento", "I")
    CurrentQuery.Cancel
    RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_NEGACAO")
    Exit Sub
  End If

  VERIFICAJAPAGO(CANCONTINUE)
  If CANCONTINUE = False Then
    bsShowMessage("Inclusão não permitida para evento já pago", "I")
    CANCONTINUE = True
    CurrentQuery.Cancel
    RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_NEGACAO")
  End If
End Sub

Public Sub TABLE_BeforePost(CANCONTINUE As Boolean)
  If alterouRevisada Then
    If(CurrentQuery.FieldByName("NEGACAOREVISADA").AsString = "S")Then
    CurrentQuery.FieldByName("USUARIOREVISOR").Value = CurrentUser
    CurrentQuery.FieldByName("DATAHORAREVISAO").Value = ServerNow
  Else
    CurrentQuery.FieldByName("USUARIOREVISOR").Clear
    CurrentQuery.FieldByName("DATAHORAREVISAO").Clear
  End If
End If



End Sub

Public Sub ChecarRefresh
  Dim qglosasOK As Object
  Set qglosasOK = NewQuery
  qglosasOK.Clear
  qglosasOK.Add("SELECT COUNT(1) QTDE           ")
  qglosasOK.Add("  FROM SAM_GUIA_EVENTOS_GLOSA  ")
  qglosasOK.Add(" WHERE GUIAEVENTO = :GUIAEVENTO")
  qglosasOK.Add("  AND GLOSAREVISADA = 'N' ")
  qglosasOK.ParamByName("GUIAEVENTO").Value = CurrentQuery.FieldByName("GUIAEVENTO").Value
  qglosasOK.Active = True
  If qglosasOK.FieldByName("QTDE").AsInteger = 0 Then
    Dim qnegacoesOK As Object
    Set qnegacoesOK = NewQuery
    qnegacoesOK.Clear
    qnegacoesOK.Add("SELECT COUNT(1) QTDE           ")
    qnegacoesOK.Add("  FROM SAM_GUIA_EVENTOS_NEGACAO")
    qnegacoesOK.Add(" WHERE GUIAEVENTO = :GUIAEVENTO")
    qnegacoesOK.Add("  AND NEGACAOREVISADA = 'N' ")
    qnegacoesOK.ParamByName("GUIAEVENTO").Value = CurrentQuery.FieldByName("GUIAEVENTO").Value
    qnegacoesOK.Active = True
    If qnegacoesOK.FieldByName("QTDE").AsInteger = 0 Then
      Dim qqtdeventosok As Object
      Set qqtdeventosok = NewQuery
      Dim vHandleGuia As Long
      qqtdeventosok.Clear
      qqtdeventosok.Add("SELECT GUIA                                 ")
      qqtdeventosok.Add("  FROM SAM_GUIA_EVENTOS E                   ")
      qqtdeventosok.Add(" WHERE E.HANDLE = :GUIAEVENTO               ")
      qqtdeventosok.ParamByName("GUIAEVENTO").Value = CurrentQuery.FieldByName("GUIAEVENTO").Value
      qqtdeventosok.Active = True
      vHandleGuia = qqtdeventosok.FieldByName("GUIA").Value
      qqtdeventosok.Clear
      qqtdeventosok.Add("SELECT COUNT(E.HANDLE) QTDE                 ")
      qqtdeventosok.Add("  FROM SAM_GUIA         G,                  ")
      qqtdeventosok.Add("       SAM_GUIA_EVENTOS E                   ")
      qqtdeventosok.Add(" WHERE G.HANDLE = :GUIA                     ")
      qqtdeventosok.Add("   AND E.GUIA = G.HANDLE                    ")
      qqtdeventosok.Add("   AND E.SITUACAO = :SITUACAO               ")
      qqtdeventosok.ParamByName("GUIA").Value = vHandleGuia
      qqtdeventosok.ParamByName("SITUACAO").AsString = vSituacaoAnteriorEvento
      qqtdeventosok.Active = True
      If qqtdeventosok.FieldByName("QTDE").AsInteger = 0 Then
        RefreshNodesWithTableInterfacePEG("SAM_GUIA")
      Else
        RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS")
      End If
      Set qqtdeventosok = Nothing
    Else
      RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_NEGACAO")
    End If
    Set qnegacoesOK = Nothing
  Else
    RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_NEGACAO")

  End If
  Set qglosasOK = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CANCONTINUE As Boolean)
	If CommandID = "BOTAORECONSIDERACAO" Then
		BOTAORECONSIDERACAO_OnClick
	End If
End Sub

Public Function RetornaFilialPEG As Long
	Dim qConsulta As Object

	Set qConsulta = NewQuery

	qConsulta.Active = False
	qConsulta.Clear
	qConsulta.Add("SELECT FILIALPROCESSAMENTO FROM SAM_PEG WHERE HANDLE = :HANDLE")
	qConsulta.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_PEG")
	qConsulta.Active = True

	RetornaFilialPEG = qConsulta.FieldByName("FILIALPROCESSAMENTO").AsInteger

	Set qConsulta = Nothing

End Function
