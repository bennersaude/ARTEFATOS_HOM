'HASH: 46590F21A77C567EAD07429351B079FE

Option Explicit
'#Uses "*ProcuraEvento"
'#Uses "*Arredonda"
'#Uses "*IsInt"
'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"
'#Uses "*VerificaPermissaoEdicaoTriagem"
'#Uses "*RecordHandleOfTableInterfacePEG"
'#Uses "*RefreshNodesWithTableInterfacePEG"

Dim vExigeComplemento As String
Dim alterouRevisada As Boolean
Dim auxGuiaEvento As Long
Dim Reconsideracao As Boolean
Dim vSituacaoAnteriorEvento As String

Public Function VERIFICAINCOMP As Boolean
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO FROM SAM_PEG WHERE HANDLE=:PEG")
  q1.ParamByName("PEG").Value = RecordHandleOfTableInterfacePEG("SAM_PEG")
  q1.Active = True
  VERIFICAINCOMP = True
  If q1.FieldByName("SITUACAO").AsString = "3" Then 'PRONTO
    Dim INTERFACE As Object
    Set INTERFACE = CreateBennerObject("SAMINCOMP.CHECK")
    INTERFACE.INICIALIZAR(CurrentSystem)
    VERIFICAINCOMP = INTERFACE.EXCLUIRINCOMPATIBILIDADES(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger)
    INTERFACE.FINALIZAR
    Set INTERFACE = Nothing
  End If
  Set q1 = Nothing
End Function

Public Sub BOTAODESCRICAOGLOSA_OnClick()
  Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "A")
  If aux = False Then
    RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  Dim INTERFACE As Object
  Set INTERFACE = CreateBennerObject("SAMPEG.PROCESSAR")
  INTERFACE.DESCRICAOGLOSA(CurrentSystem, CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger)
  Set INTERFACE = Nothing

End Sub

Public Sub MOTIVOGLOSA_OnExit()
  If CurrentQuery.State <>1 Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Clear
    SQL.Add("SELECT GLOSADEPENDENTE FROM SAM_MOTIVOGLOSA WHERE HANDLE=:MOTIVOGLOSA")
    SQL.ParamByName("MOTIVOGLOSA").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    SQL.Active = True
    CurrentQuery.FieldByName("GLOSADEPENDENTE").Value = SQL.FieldByName("GLOSADEPENDENTE").AsString
    Set SQL = Nothing
  End If

End Sub

Public Sub MOTIVOGLOSA_OnPopup(ShowPopup As Boolean)

  Dim RestringeGlosaAutomatica As String
  RestringeGlosaAutomatica = " AND SAM_MOTIVOGLOSA.HANDLE NOT IN (SELECT SAM_MOTIVOGLOSA_SIS.SAMMOTIVOGLOSA FROM SAM_MOTIVOGLOSA_SIS)"

  Dim SQL As Object
  Set SQL = NewQuery
  On Error GoTo prox
  SQL.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA WHERE CODIGOGLOSA=:CODIGOGLOSA" + RestringeGlosaAutomatica)
  SQL.ParamByName("CODIGOGLOSA").AsString = MOTIVOGLOSA.Text
  SQL.Active = True
  If Not SQL.EOF Then
    CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger = SQL.FieldByName("HANDLE").AsInteger
    ShowPopup = False
    HabilitaCamposGlosa
    Set SQL = Nothing
    Exit Sub
  End If
  Set SQL = Nothing
prox:

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim INTERFACE As Object

  ShowPopup = False

  vColunas = "CODIGOGLOSA|SAM_MOTIVOGLOSA.DESCRICAO|SAM_TIPOMOTIVOGLOSA.DESCRICAO"
  vCampos = "Código|Descrição|Tipo Glosa"
  vCriterio = "SAM_MOTIVOGLOSA.ATIVA='S'"
  vCriterio = vCriterio + RestringeGlosaAutomatica

  Set INTERFACE = CreateBennerObject("Procura.Procurar")
  vHandle = INTERFACE.Exec(CurrentSystem, "SAM_MOTIVOGLOSA|SAM_TIPOMOTIVOGLOSA[SAM_MOTIVOGLOSA.TIPOMOTIVOGLOSA=SAM_TIPOMOTIVOGLOSA.HANDLE]", vColunas, 2, vCampos, vCriterio, "Tabela de Motivo de Glosas", True, "")

  Set INTERFACE = Nothing

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("MOTIVOGLOSA").Value = vHandle
    HabilitaCamposGlosa
  End If

End Sub

Public Sub QTDGLOSA_OnExit()
  Dim vReadOnlyOriginal As Boolean
  vReadOnlyOriginal = VALORGLOSA.ReadOnly

  If (CurrentQuery.State = 3) Then
    If CurrentQuery.FieldByName("QTDGLOSA").AsFloat > 0 Then 'Alterado de AsInteger para AsFloat na SMS 67809 - 02.09.2006
      VALORGLOSA.ReadOnly = True
      QTDRECONSIDERADA.ReadOnly = False
      VALORRECONSIDERADO.ReadOnly = True

    Else
      VALORGLOSA.ReadOnly = False
      VALORRECONSIDERADO.ReadOnly = False
      If vReadOnlyOriginal Then
        VALORGLOSA.SetFocus
      End If
    End If
  End If

  If (VALORGLOSA.ReadOnly) And (CurrentQuery.State = 3 Or CurrentQuery.State = 2) Then

    Dim q1 As Object
    Set q1 = NewQuery

    q1.Clear
    q1.Add("SELECT QTDAPRESENTADA FROM SAM_GUIA_EVENTOS WHERE HANDLE=" + CurrentQuery.FieldByName("GUIAEVENTO").AsString)
    q1.Active = True
    If Arredonda(CurrentQuery.FieldByName("QTDGLOSA").AsFloat) > Arredonda(q1.FieldByName("QTDAPRESENTADA").AsFloat) Then
       CurrentQuery.FieldByName("QTDGLOSA").AsFloat = Arredonda(q1.FieldByName("QTDAPRESENTADA").AsFloat)
    End If

    Dim INTERFACE As Object
    Set INTERFACE = CreateBennerObject("sampeg.Rotinas")
    CurrentQuery.FieldByName("VALORGLOSA").Value = INTERFACE.SugerirValorGlosa(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, CurrentQuery.FieldByName("QTDGLOSA").AsFloat)
    Set INTERFACE = Nothing
  End If

End Sub

Public Sub QTDRECONSIDERADA_OnExit()
  Dim vQtdeReconsiderada As Double
  Dim vQtdeGlosa         As Double

  Dim q1 As Object
  Set q1 = NewQuery


  q1.Clear
  q1.Add("SELECT VALORCALCPAGTO, QTDAPRESENTADA FROM SAM_GUIA_EVENTOS WHERE HANDLE=" + CurrentQuery.FieldByName("GUIAEVENTO").AsString)
  q1.Active = True

  vQtdeReconsiderada  = arredonda(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat)
  vQtdeGlosa          = arredonda(CurrentQuery.FieldByName("QTDGLOSA").AsFloat)

  If((CurrentQuery.State = 2 Or CurrentQuery.State = 3 Or CurrentQuery.State = 2 Or CurrentQuery.State = 2) And _
     arredonda(vQtdeReconsiderada) > 0 And arredonda(vQtdeGlosa) > 0) Then

    Dim INTERFACE  As Object
    Set INTERFACE=CreateBennerObject("sampeg.processar")
    CurrentQuery.FieldByName("VALORRECONSIDERADO").Value = INTERFACE.SugerirValorReconsiderado(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat)
    Set INTERFACE=Nothing

  Else
    If (CurrentQuery.State=2 And _
        arredonda(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat) > 0 And _
        arredonda(CurrentQuery.FieldByName("QTDGLOSA").AsFloat) = 0) Then
      CurrentQuery.FieldByName("QTDRECONSIDERADA").Value = CurrentQuery.FieldByName("QTDGLOSA").Value
    End If
  End If

  If (CurrentQuery.State = 3 Or CurrentQuery.State = 2) Then
    If (CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0) And _
       (CurrentQuery.FieldByName("QTDGLOSA").AsFloat > 0) Then
     CurrentQuery.FieldByName("VALORRECONSIDERADO").Value = 0
    End If
  End If


q1.Active = False
End Sub



Public Sub TABLE_AfterCancel()
  SessionVar("Reconsiderar") = "N"
End Sub

Public Sub TABLE_AfterDelete()
  Dim INTERFACE As Object
  Set INTERFACE = CreateBennerObject("BSPro000.Rotinas")
  INTERFACE.RevisarEvento(CurrentSystem, auxGuiaEvento, "TOTAL", True)
  Set INTERFACE = Nothing
End Sub

Public Sub TABLE_AfterInsert()


  COMPLEMENTORECONSIDERACAO.ReadOnly = False
  MOTIVORECONSIDERACAO.ReadOnly      = False
  VALORRECONSIDERADO.ReadOnly        = True
  QTDRECONSIDERADA.ReadOnly          = True

  If WebMode Then
    'Em modo Web deve permtir digitar os campos de quantidade e valor de glosa
    'A permissão para se informar a quantidade ou o valor deve ser feita no BeforePost
      QTDGLOSA.ReadOnly   = False
      VALORGLOSA.ReadOnly = False
  Else
    If CurrentQuery.FieldByName("ORIGEMGLOSA").AsInteger = 1 Then
    ' - só vai se saber que campos habilitar após selecionado o tipo de glosa (de valor ou de qtd)
      QTDRECONSIDERADA.ReadOnly   = True
      VALORRECONSIDERADO.ReadOnly = True
      QTDGLOSA.ReadOnly           = True
      VALORGLOSA.ReadOnly         = True
    End If
  End If
End Sub

Public Sub TABLE_AfterCommitted()

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
        RefreshNodesWithTable("SAM_GUIA")
      Else
        RefreshNodesWithTable("SAM_GUIA_EVENTOS")
      End If
      Set qqtdeventosok = Nothing
    End If
    Set qnegacoesOK = Nothing
  End If
  Set qglosasOK = Nothing
End Sub


Public Sub TABLE_AfterPost()

  Dim INTERFACE As Object

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT COUNT(1) GLOSAVALOR      ")
  SQL.Add("FROM SAM_GUIA_EVENTOS_GLOSA     ")
  SQL.Add("WHERE GUIAEVENTO = :GUIAEVENTO  ")
  SQL.Add(" AND QTDRECONSIDERADA = 0       ")
  SQL.Add(" AND VALORRECONSIDERADO <> 0    ")
  SQL.ParamByName("GUIAEVENTO").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
  SQL.Active = True

  If (SQL.FieldByName("GLOSAVALOR").AsInteger > 0 And _
	 CurrentQuery.FieldByName("QTDGLOSA").AsInteger > 0) Then

    Set INTERFACE = CreateBennerObject("BSPro000.Rotinas")

	INTERFACE.ReprocessarEvento(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, "TOTAL", True)

	bsShowMessage("Glosas reprocessadas!","I")

  End If

  Set SQL = Nothing

  Set INTERFACE=CreateBennerObject("BSPro000.Rotinas")

  INTERFACE.RevisarEvento(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, "REVISAR", True)

  Set INTERFACE = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  PreencheSessionVar

  BOTAORECONSIDERACAO.Enabled = VerificarPermissaoUsuarioPegTriado(False)

  alterouRevisada = False

  Dim vsMensagem As String

  Dim qGlosa As Object
  Set qGlosa = NewQuery
  qGlosa.Add("SELECT DESCRICAOGLOSA FROM SAM_MOTIVOGLOSA WHERE HANDLE=:HANDLE")
  qGlosa.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  qGlosa.Active = True

  Dim vsAux As String
  Dim vsAux2 As String
  Dim viCount As Integer

  vsAux = qGlosa.FieldByName("DESCRICAOGLOSA").AsString

  If WebMode Then
  	DESCRICAOGLOSA.Text = "Descrição da glosa no sistema: " + vsAux
  	DESCRICAOGLOSA2.Text = ""

	  Dim vCriterio As String

	  vCriterio = "A.ATIVA='S'"

	  vCriterio = vCriterio + "AND A.HANDLE NOT IN ("
	  vCriterio = vCriterio + "SELECT SAM_MOTIVOGLOSA_SIS.SAMMOTIVOGLOSA FROM SAM_MOTIVOGLOSA_SIS)"

	  MOTIVOGLOSA.WebLocalWhere = vCriterio

  ElseIf VisibleMode Then
	  BOTAORECONSIDERACAO2.Visible = False

	  If InStr(vsAux, Chr(13)) Then
	    viCount = InStr(vsAux, Chr(13))
	  	vsAux2 = Mid(vsAux, 1 , viCount - 1)
	  	vsAux = Mid(vsAux, viCount + 2, Len(vsAux))
	  End If


      DESCRICAOGLOSA.Text = "Descrição da glosa no sistema: " + vsAux2
	  DESCRICAOGLOSA2.Text = vsAux
  End If

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DECISAO, APLICARECURSO FROM SAM_MOTIVOGLOSA WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  SQL.Active = True
  If SQL.FieldByName("DECISAO").AsString = "A" Then
    ROTULODECISAO.Text = "Decisão: Administrativo"
  ElseIf SQL.FieldByName("DECISAO").AsString = "M" Then
    ROTULODECISAO.Text = "Decisão: Médico"
  ElseIf SQL.FieldByName("DECISAO").AsString = "O" Then
    ROTULODECISAO.Text = "Decisão: Odontológico"
  Else
    ROTULODECISAO.Text = "Decisão: ------- "
  End If

  If SQL.FieldByName("APLICARECURSO").AsString = "S" Then
    ROTULORECURSO.Text = "Aplica Recurso"
  Else
    ROTULORECURSO.Text = "Não Aplica Recurso"
  End If
  Dim qSituacaoEvento As Object
  Set qSituacaoEvento = NewQuery
  If CurrentQuery.FieldByName("GUIAEVENTO").AsInteger > 0 Then
    qSituacaoEvento.Clear
    qSituacaoEvento.Add("SELECT SITUACAO FROM SAM_GUIA_EVENTOS WHERE HANDLE = :HANDLE")
    qSituacaoEvento.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("GUIAEVENTO").Value
    qSituacaoEvento.Active = True
    vSituacaoAnteriorEvento = qSituacaoEvento.FieldByName("SITUACAO").AsString
  Else
    qSituacaoEvento.Clear
    qSituacaoEvento.Add("SELECT SITUACAO FROM SAM_GUIA_EVENTOS WHERE HANDLE = :HANDLE")
    qSituacaoEvento.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_GUIA_EVENTOS")
    qSituacaoEvento.Active = True
    vSituacaoAnteriorEvento = qSituacaoEvento.FieldByName("SITUACAO").AsString
  End If
  If CurrentQuery.State = 1 Then
    If vSituacaoAnteriorEvento = "8" Then
      TableReadOnly = True
    Else
      TableReadOnly = False
    End If
  End If

  If Not CurrentQuery.FieldByName("ORIGEMGLOSA").IsNull Then
    If CurrentQuery.FieldByName("ORIGEMGLOSA").AsInteger = 3 Then
  	  VALORGLOSA.ReadOnly = True
    End If
  End If

  MOTIVORECONSIDERACAO.ReadOnly      = True
  VALORRECONSIDERADO.ReadOnly        = True
  QTDRECONSIDERADA.ReadOnly          = True

  If WebMode And _
     CurrentQuery.State = 3 Then
    QTDGLOSA.ReadOnly   = False
  Else
    QTDGLOSA.ReadOnly = True
  End If

  If PermiteReconsiderar(vsMensagem) = False Then
    VALORGLOSA.ReadOnly = True
    COMPLEMENTORECONSIDERACAO.ReadOnly = True
    Exit Sub
  Else
    VALORGLOSA.ReadOnly = False
    COMPLEMENTORECONSIDERACAO.ReadOnly = False
  End If

  Set qSituacaoEvento = Nothing
  SQL.Active = False
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeCancel(CanContinue As Boolean)
  If MOTIVORECONSIDERACAO.ReadOnly = False Then
    COMPLEMENTORECONSIDERACAO.ReadOnly = True
    MOTIVORECONSIDERACAO.ReadOnly      = True
    VALORRECONSIDERADO.ReadOnly        = True
    QTDRECONSIDERADA.ReadOnly          = True
    QTDGLOSA.ReadOnly                  = True
    VALORGLOSA.ReadOnly                = True
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  PreencheSessionVar

  If Not VerificarPermissaoUsuarioPegTriado(False) Then
   	bsShowMessage("Alteração não permitida para este usuário.", "I")
 	CanContinue = False
	Exit Sub
  End If

  CanContinue = CheckFilialProcessamento(CurrentSystem, RetornaFilialPEG, "E")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_GLOSA")
    Exit Sub
  End If

  Dim agrupadorFechado As Boolean

  agrupadorFechado = VerificaAgrupadorPagamentoFechado

  If (agrupadorFechado) Then
  	BsShowMessage("Não é permitida a alteração de glosa ligada à registro de pagamento fechado.","E")
    CanContinue = False
    Exit Sub
  End If

  VERIFICAJAPAGO(CanContinue)
  If CanContinue = False Then
    bsShowMessage("Alteração não permitida nesta fase", "I")
    Exit Sub
  Else
    If CurrentQuery.FieldByName("ORIGEMGLOSA").AsInteger <>1 Then
      bsShowMessage("Só é permitida a exclusão de glosas manuais", "E")
      CanContinue = False
      auxGuiaEvento = 0
    Else
      auxGuiaEvento = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
    End If
  End If
End Sub


Public Sub VERIFICAJAPAGO(CanContinue As Boolean)
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
    CanContinue = False
  Else
    CanContinue = True
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  PreencheSessionVar

  If Not VerificarPermissaoUsuarioPegTriado(False) Then
   	bsShowMessage("Alteração não permitida para este usuário.", "I")
 	CanContinue = False
 	RecordReadOnly = Not CanContinue
 	Exit Sub
  End If

  CanContinue = CheckFilialProcessamento(CurrentSystem, RetornaFilialPEG, "I")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_GLOSA")
    Exit Sub
  End If

  VERIFICAJAPAGO(CanContinue)
  If CanContinue = False Then
    bsShowMessage("Inclusão não permitida nesta fase", "E")
    CanContinue = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim valorapresentado As Variant
  Dim valorcalculado As Variant
  Dim vComoPagar As String
  Dim valorGlosado As Currency
  Dim ValorUnitarioGlosa As Double
  Dim MaxRecon As Double
  Dim vRecebedor As Long

  Dim sqlev As Object
  Dim SQLAUX As Object
  Dim SQL As Object

  On Error GoTo lbSair

  Dim agrupadorFechado As Boolean

  agrupadorFechado = VerificaAgrupadorPagamentoFechado

  If (agrupadorFechado) Then
  	BsShowMessage("Não é permitida a alteração de glosa ligada à registro de pagamento fechado.","E")
    GoTo lbSairComErro
  End If

  CriaTabelaTemporariaSQLServer
  QTDRECONSIDERADA_OnExit

  'SMS 41587
  If CurrentQuery.FieldByName("QTDGLOSA").AsFloat > 0 Then
    CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = "S"
  Else
    CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = "N"
  End If


  If(CurrentQuery.FieldByName("VALORGLOSA").AsFloat = 0)And(CurrentQuery.FieldByName("QTDGLOSA").AsFloat = 0)And(CurrentQuery.FieldByName("ORIGEMGLOSA").AsInteger = 1)Then
    BsShowMessage("Favor informar valor ou quantidade glosada","E")
    GoTo lbSairComErro
  End If

  Set SQL = NewQuery

  If WebMode And _
     CurrentQuery.State = 3 Then
    'Em modo Web, em virtude de não se ter eventos de campos para setar o readonly
    'após informar o tipo da glosa, o tratamento deve ser feito no post do registro
    SQL.Clear
    SQL.Add("SELECT GLOSADEPENDENTE FROM SAM_MOTIVOGLOSA WHERE HANDLE=:MOTIVOGLOSA")
    SQL.ParamByName("MOTIVOGLOSA").Value=CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    SQL.Active=True

    CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = SQL.FieldByName("GLOSADEPENDENTE").AsString

    If SQL.FieldByName("GLOSADEPENDENTE").AsString = "S" Then
      If CurrentQuery.FieldByName("VALORGLOSA").AsFloat > 0 Then
        Set SQL = Nothing
        bsShowMessage("Motivo de glosa ""dependente"" não permite que se informe valor glosado! Informar somente a quantidade a ser glosada.", "E")
        CanContinue = False
        Exit Sub
      End If
    Else
      If CurrentQuery.FieldByName("QTDGLOSA").AsFloat > 0 Then
        Set SQL = Nothing
        bsShowMessage("Motivo de glosa não permite que se informe quantidade glosada! Informar somente o valor a ser glosado.", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  End If

  If(CurrentQuery.FieldByName("GLOSAREVISADA").AsString = "S")Then
    CurrentQuery.FieldByName("USUARIOREVISOR").Value = CurrentUser
    CurrentQuery.FieldByName("DATAHORAREVISAO").Value = ServerNow
  Else
    CurrentQuery.FieldByName("USUARIOREVISOR").Clear
    CurrentQuery.FieldByName("DATAHORAREVISAO").Clear
    'se nao está revisada, nunca pode estar reconsiderada
    CurrentQuery.FieldByName("DATAHORALIBERACAO").Clear
    CurrentQuery.FieldByName("USUARIOLIBERACAO").Clear
  End If


  SQL.Clear
  SQL.Add("SELECT ATIVA FROM SAM_MOTIVOGLOSA WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  SQL.Active = True
  If SQL.FieldByName("ATIVA").AsString = "N" Then
    BsShowMessage("O motivo de glosa está desativado","E")
    GoTo lbSairComErro
  End If

  SQL.Clear
  SQL.Add("SELECT QTDAPRESENTADA, VALORAPRESENTADO, VALORCALCPAGTO FROM SAM_GUIA_EVENTOS WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTableInterfacePEG("SAM_GUIA_EVENTOS")
  SQL.Active = True

  valorcalculado = SQL.FieldByName("VALORCALCPAGTO").AsFloat
  valorapresentado = SQL.FieldByName("VALORAPRESENTADO").Value

  If SQL.FieldByName("QTDAPRESENTADA").AsFloat < CurrentQuery.FieldByName("QTDGLOSA").AsFloat Then
    BsShowMessage("A quantidade glosada não pode ser superior à quantidade apresentada","E")
    GoTo lbSairComErro
  End If

  If CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat <0 Then
    BsShowMessage("O valor reconsiderado deve ser maior que zero","E")
    GoTo lbSairComErro
  End If

  If CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat <0 Then
    BsShowMessage("O quantidade reconsiderada deve ser maior que zero","E")
    GoTo lbSairComErro
  End If

  If CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = "N" Then

	If (CurrentQuery.FieldByName("ORIGEMGLOSA").AsString ="1") Then

		Dim vMaxGlosar As Currency

		Dim vDllBSPro006 As Object

		Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventosGlosa")

		vMaxGlosar = vDllBSPro006.ValorMaximoParaGlosar(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger)

		Set vDllBSPro006 = Nothing

		If CurrentQuery.FieldByName("VALORGLOSA").AsCurrency > vMaxGlosar Then
			BsShowMessage("Valor máximo para glosar: " + Format(vMaxGlosar, "##0.00"),"E")
			GoTo lbSairComErro
		End If
    End If

    Set sqlev = NewQuery
    sqlev.Add("SELECT VALORGLOSADO, QTDAPRESENTADA, QTDGLOSADA FROM SAM_GUIA_EVENTOS WHERE HANDLE=" + Str(RecordHandleOfTableInterfacePEG("SAM_GUIA_EVENTOS")))
    sqlev.Active = True

    ValorUnitarioGlosa = CurrentQuery.FieldByName("VALORGLOSA").AsFloat / sqlev.FieldByName("QTDAPRESENTADA").AsFloat
    MaxRecon = ValorUnitarioGlosa * (sqlev.FieldByName("QTDAPRESENTADA").AsFloat - sqlev.FieldByName("QTDGLOSADA").AsFloat)

    If Round(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat,2) > Round(MaxRecon,2) Then
      BsShowMessage("Máximo para reconsiderar: " + Str(MaxRecon) + ". Para reconsiderar " + Str(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat) + ", deve-se primeiro reconsiderar As glosas dependentes", "E")
	  GoTo lbSairComErro
    End If

  End If
  If(arredonda(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat)>arredonda(CurrentQuery.FieldByName("QTDGLOSA").AsFloat))Or _
     (arredonda(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat)>arredonda(CurrentQuery.FieldByName("VALORGLOSA").AsFloat))Then
    BsShowMessage("A reconsideração não pode ser maior que a glosa","E")
    GoTo lbSairComErro
  End If

  If(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat >0) Or (CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat >0) Then
     If CurrentQuery.FieldByName("GLOSAREVISADA").AsString = "S" Then
       CurrentQuery.FieldByName("DATAHORALIBERACAO").Value = ServerNow
       CurrentQuery.FieldByName("USUARIOLIBERACAO").Value = CurrentUser
     Else
      CurrentQuery.FieldByName("DATAHORALIBERACAO").Clear
      CurrentQuery.FieldByName("USUARIOLIBERACAO").Clear
      CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
      CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
     End If
  Else
    'marcar o usuário revisao
    CurrentQuery.FieldByName("DATAHORALIBERACAO").Clear
    CurrentQuery.FieldByName("USUARIOLIBERACAO").Clear
    CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
    CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0

  End If

  If CurrentQuery.FieldByName("GLOSAREVISADA").AsString = "S" Then
    SQL.Clear
    SQL.Add("SELECT EXIGECOMPLEMENTO FROM SAM_MOTIVOGLOSA WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    SQL.Active = True
    vExigeComplemento = SQL.FieldByName("EXIGECOMPLEMENTO").AsString
    If vExigeComplemento = "S" Then
      If CurrentQuery.FieldByName("COMPLEMENTO").IsNull Or _
        Trim(CurrentQuery.FieldByName("COMPLEMENTO").AsString) = "" Then
        BsShowMessage("Motivo de glosa exige a digitação do complemento!","E")
        GoTo lbSairComErro
      End If
    End If
  End If

  If CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat <>0 And CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat <>0 Then
    If Not CurrentQuery.FieldByName("MOTIVORECONSIDERACAO").IsNull Then 'Henrique SMS: 14647
      SQL.Clear
      SQL.Active = False
      SQL.Add("SELECT EXIGECOMPLEMENTO FROM SAM_MOTIVORECONSIDERACAO WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MOTIVORECONSIDERACAO").AsInteger
      SQL.Active = True
      If SQL.FieldByName("EXIGECOMPLEMENTO").AsString = "S" Then
        If CurrentQuery.FieldByName("COMPLEMENTORECONSIDERACAO").IsNull Or _
          Trim(CurrentQuery.FieldByName("COMPLEMENTORECONSIDERACAO").AsString) = "" Then
          BsShowMessage("Motivo de reconsideração exige a digitação do complemento reconsideração!","E")
          GoTo lbSairComErro
        End If
      End If
    End If
    If MOTIVORECONSIDERACAO.ReadOnly = False Then
      COMPLEMENTORECONSIDERACAO.ReadOnly = True
      MOTIVORECONSIDERACAO.ReadOnly      = True
      VALORRECONSIDERADO.ReadOnly        = True
      QTDRECONSIDERADA.ReadOnly          = True
    End If
    SQL.Active = False
  End If

  CurrentQuery.FieldByName("VALORGLOSA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("VALORGLOSA").AsFloat, "##0.00"))
  CurrentQuery.FieldByName("QTDGLOSA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("QTDGLOSA").AsFloat, "##0.00"))
  CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat, "##0.00"))
  CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = CDbl(Format(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat, "##0.00"))

  GoTo lbFinalizar

  lbSair:
    BsShowMessage("Erro: " + Err.Description, "E")

  lbSairComErro:
    CanContinue = False

  lbFinalizar:
    Set SQL = Nothing
    Set SQLAUX = Nothing
    Set sqlev = Nothing

End Sub

Public Function VerificaReconsideracao As String
  Dim q1 As Object
  Dim Reconsiderar As String
  Dim vNaoReconsiderarPara As String

  Set q1=NewQuery

  VerificaReconsideracao = ""
  q1.Clear
  q1.Add("SELECT M.HANDLE ")
  q1.Add("  FROM SAM_MOTIVOGLOSA_SIS M, ")
  q1.Add("       SAM_MOTIVOGLOSA SAM, ")
  q1.Add("       SIS_MOTIVOGLOSA SIS ")
  q1.Add("WHERE SAM.HANDLE=:MOTIVOGLOSA")
  q1.Add("  AND (SIS.CODIGO IN (6010,7000,8000,8010,9650))")
  q1.Add("  AND SAM.HANDLE=M.SAMMOTIVOGLOSA")
  q1.Add("  AND SIS.HANDLE=M.SISMOTIVOGLOSA")
  q1.ParamByName("MOTIVOGLOSA").Value=CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  q1.Active=True
  If Not q1.EOF Then
    VerificaReconsideracao = "O Motivo da glosa não suporta reconsideração"
    Set q1=Nothing
    Exit Function
  End If
  Set q1=Nothing

  'Verificar se o motivo de glosa está inserido na carga abaixo do usuário, se não estiver então nào será possível reconsiderar a glosa
  Dim qUsuarioGlosa As Object
  Set qUsuarioGlosa = NewQuery

  qUsuarioGlosa.Clear
  qUsuarioGlosa.Add("SELECT EXIGERESPONSAVEL FROM SAM_MOTIVOGLOSA WHERE HANDLE = :HMOTIVOGLOSA")
  qUsuarioGlosa.ParamByName("HMOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  qUsuarioGlosa.Active = True

  If qUsuarioGlosa.FieldByName("EXIGERESPONSAVEL").AsString = "S" Then

    Dim vQtdeUsuarioGlosa As Integer

    qUsuarioGlosa.Clear
    qUsuarioGlosa.Add("SELECT COUNT(1) QTDE")
    qUsuarioGlosa.Add("  FROM SAM_USUARIO_MOTIVOGLOSA")
    qUsuarioGlosa.Add(" WHERE USUARIO = :HUSUARIO")
    qUsuarioGlosa.Add("   AND MOTIVOGLOSA = :HMOTIVOGLOSA")
    qUsuarioGlosa.ParamByName("HMOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    qUsuarioGlosa.ParamByName("HUSUARIO").AsInteger = CurrentUser
    qUsuarioGlosa.Active = True

    vQtdeUsuarioGlosa = qUsuarioGlosa.FieldByName("QTDE").AsInteger

    If vQtdeUsuarioGlosa = 0 Then
      'Agora será verificado se esta glosa está cadastrada no grupo de segurança principal do usuário corrente
      qUsuarioGlosa.Clear
      qUsuarioGlosa.Add("SELECT SUM(X.QTDE) QTDE")
      qUsuarioGlosa.Add("  FROM (")
      qUsuarioGlosa.Add("        SELECT COUNT(1) QTDE ")
      qUsuarioGlosa.Add("          FROM SAM_GRUPO_MOTIVOGLOSA   GM, ")
      qUsuarioGlosa.Add("               Z_GRUPOUSUARIOS          U")
      qUsuarioGlosa.Add("         WHERE GM.GRUPO = U.GRUPO ")
      qUsuarioGlosa.Add("           AND GM.MOTIVOGLOSA = :HMOTIVOGLOSA")
      qUsuarioGlosa.Add("           AND U.HANDLE = :HUSUARIO")
      qUsuarioGlosa.Add("         UNION ALL")
      qUsuarioGlosa.Add("        SELECT COUNT(1) QTDE ")
      qUsuarioGlosa.Add("          FROM SAM_GRUPO_MOTIVOGLOSA   GM, ")
      qUsuarioGlosa.Add("               Z_GRUPOUSUARIOGRUPOS    UG	")
      qUsuarioGlosa.Add("         WHERE UG.GRUPOADICIONADO = GM.GRUPO")
      qUsuarioGlosa.Add("           AND GM.MOTIVOGLOSA = :HMOTIVOGLOSA")
      qUsuarioGlosa.Add("           AND UG.USUARIO = :HUSUARIO")
      qUsuarioGlosa.Add("       ) X")
      qUsuarioGlosa.ParamByName("HMOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
      qUsuarioGlosa.ParamByName("HUSUARIO").AsInteger = CurrentUser
      qUsuarioGlosa.Active = True

      vQtdeUsuarioGlosa = qUsuarioGlosa.FieldByName("QTDE").AsInteger
      Set qUsuarioGlosa = Nothing

      If vQtdeUsuarioGlosa = 0 Then
         VerificaReconsideracao = "Usuário não tem permissão para reconsiderar a glosa."

        Exit Function
      End If
    Else
      Set qUsuarioGlosa = Nothing
    End If
  Else
    Set qUsuarioGlosa = Nothing
  End If

End Function

Public Function HabilitaCamposGlosa As Integer

    Dim SQL As Object
    Set SQL=NewQuery
    SQL.Clear
    SQL.Add("SELECT GLOSADEPENDENTE FROM SAM_MOTIVOGLOSA WHERE HANDLE=:MOTIVOGLOSA")
    SQL.ParamByName("MOTIVOGLOSA").Value=CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    SQL.Active=True
    CurrentQuery.FieldByName("GLOSADEPENDENTE").Value=SQL.FieldByName("GLOSADEPENDENTE").AsString
    If SQL.FieldByName("GLOSADEPENDENTE").AsString = "S" Then
      CurrentQuery.FieldByName("VALORGLOSA").AsFloat         = 0
      CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
      VALORGLOSA.ReadOnly         = True
      VALORRECONSIDERADO.ReadOnly = True
      CurrentQuery.FieldByName("QTDGLOSA").AsFloat         = 0
	  CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
      QTDGLOSA.ReadOnly         = False
      QTDRECONSIDERADA.ReadOnly = False
    Else
      CurrentQuery.FieldByName("QTDGLOSA").AsFloat         = 0
      CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
      QTDGLOSA.ReadOnly         = True
      QTDRECONSIDERADA.ReadOnly = True
      CurrentQuery.FieldByName("VALORGLOSA").AsFloat         = 0
	  CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
      VALORGLOSA.ReadOnly         = False
      VALORRECONSIDERADO.ReadOnly = False
    End If

    Set SQL=Nothing
End Function

Public Function PermiteReconsiderar(psMensagem As String) As Boolean

  Dim aux       As Boolean
  Dim vbResult  As Boolean
  Dim SQL2      As Object
  Dim SQL       As Object

  psMensagem = ""
  vbResult = True

  On Error GoTo lbErro:

  Set SQL = NewQuery
  SQL.Add("SELECT FILIALPROCESSAMENTO ")
  SQL.Add("  FROM SAM_GUIA_EVENTOS ")
  SQL.Add(" WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
  SQL.Active = True

  psMensagem = VerificaReconsideracao
  If Len(psMensagem) > 0 Then
    vbResult = False
    GoTo lbFim
  End If

  Reconsideracao = True

  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, SQL.FieldByName("FILIALPROCESSAMENTO").AsInteger, "A")
  If aux = False Then
    RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_GLOSA")
    psMensagem = "Usuário sem permissão na filial!"
    vbResult = False
    GoTo lbFim
  End If

  SQL.Clear
  SQL.Add("SELECT COUNT(GE.HANDLE) TOTAL")
  SQL.Add("  FROM SAM_GUIA_EVENTOS        GE, ")
  SQL.Add("       SAM_GUIA_EVENTOS_ACERTO GEA")
  SQL.Add(" WHERE GEA.GUIAEVENTO    = GE.HANDLE")
  SQL.Add("   AND GEA.TABTIPOACERTO = 1")
  SQL.Add("   And GE.HANDLE         = " + Str(CurrentQuery.FieldByName("GUIAEVENTO").AsInteger))
  SQL.Active = True

  If SQL.FieldByName("TOTAL").AsInteger <>0 Then
    psMensagem = "Reconsideração não permitida: Evento sofreu movimento de acerto do tipo alteração!"
    vbResult = False
    GoTo lbFim
  End If

  Set SQL2 = NewQuery
  SQL2.Add("SELECT COUNT(GE.HANDLE) TOTAL")
  SQL2.Add("  FROM SAM_GUIA_EVENTOS        GE, ")
  SQL2.Add("       SAM_GUIA_EVENTOS_ACERTO GEA")
  SQL2.Add(" WHERE GEA.GUIAEVENTO      = GE.HANDLE")
  SQL2.Add("   AND GEA.ACERTOREALIZADO = 'N'")
  SQL2.Add("   And GE.HANDLE           = " + Str(CurrentQuery.FieldByName("GUIAEVENTO").AsInteger))
  SQL2.Active = True

  If SQL2.FieldByName("TOTAL").AsInteger <>0 Then
    psMensagem = "Reconsideração não permitida: Existe movimento de acerto não faturado para este Evento!"
    vbResult = False
    GoTo lbFim
  End If

  GoTo lbFim

  lbErro:
    vbResult = False
    psMensagem = "Erro: " + Err.Description

  lbFim:
    If vbResult Then
      PermiteReconsiderar = True
    Else
      PermiteReconsiderar = False
    End If

    Set SQL       = Nothing
    Set SQL2      = Nothing

End Function

Public Sub ReconsiderarGlosa()

  QTDGLOSA.ReadOnly   = True
  VALORGLOSA.ReadOnly = True

  Dim Q         As Object
  Dim INTERFACE As Object

  On Error GoTo lbErro:

  Set Q = NewQuery
  Q.Clear
  Q.Add("SELECT HANDLE,FATURAPAGAMENTO ")
  Q.Add("  FROM SAM_GUIA_EVENTOS ")
  Q.Add(" WHERE HANDLE = :HGUIAEVENTO")
  Q.ParamByName("HGUIAEVENTO").Value = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
  Q.Active = True

  If Q.FieldByName("FATURAPAGAMENTO").IsNull Then
    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If

    If CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = "S" Then
      QTDRECONSIDERADA.ReadOnly   = False
      VALORRECONSIDERADO.ReadOnly = True
    Else
      VALORRECONSIDERADO.ReadOnly = False
      QTDRECONSIDERADA.ReadOnly   = True
    End If

    COMPLEMENTORECONSIDERACAO.ReadOnly = False
    MOTIVORECONSIDERACAO.ReadOnly      = False

    CurrentQuery.FieldByName("VALORRECONSIDERADO").Value = CurrentQuery.FieldByName("VALORGLOSA").AsFloat
    CurrentQuery.FieldByName("QTDRECONSIDERADA").Value = CurrentQuery.FieldByName("QTDGLOSA").AsFloat

    alterouRevisada = True
    CurrentQuery.FieldByName("GLOSAREVISADA").AsString = "S"

    QTDRECONSIDERADA.SetFocus 'Não retirar, necessário no sistema Desktop

    If WebMode Then
      QTDRECONSIDERADA_OnExit
    End If

  End If

  GoTo lbFim

  lbErro:
    BsShowMessage("Erro: " + Err.Description, "E")

  lbFim:
	Set Q         = Nothing
    Set interface = Nothing

End Sub


Public Sub BOTAORECONSIDERACAO_OnClick()

  SessionVar("Reconsiderar") = "S"

  If vSituacaoAnteriorEvento <> "8" Then

	  If CurrentQuery.State <>1 Then
	    BsShowMessage("O registro não pode estar em edição", "I")
	    Exit Sub
	  End If

      Dim vsMensagem As String

	  If PermiteReconsiderar(vsMensagem) = True Then
	    ReconsiderarGlosa
	  Else
	    bsShowMessage(vsMensagem, "I")
	  End If

  End If
End Sub

Public Sub TABLE_BeforeScroll()
  SessionVar("Reconsiderar") = "N"
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

  SessionVar("Reconsiderar") = "N"
  If CommandID = "BOTAORECONSIDERACAO" Then
    SessionVar("Reconsiderar") = "S"
  Else
	CurrentQuery.FieldByName("QTDRECONSIDERADA").ReadOnly = False
	CurrentQuery.FieldByName("VALORRECONSIDERADO").ReadOnly = False

    CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
    CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  PreencheSessionVar

  If Not VerificarPermissaoUsuarioPegTriado(False) Then
   	bsShowMessage("Alteração não permitida para este usuário.", "I")
 	CanContinue = False
 	Exit Sub
  End If
  Dim vsMensagem As String

  If SessionVar("Reconsiderar") = "S" Then
    If PermiteReconsiderar(vsMensagem) = False Then
      bsShowMessage(vsMensagem, "E")
      CanContinue = False
      Exit Sub
    End If
  Else
    SessionVar("Reconsiderar") = "N"
  End If

  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTableInterfacePEG("SAM_GUIA_EVENTOS_GLOSA")
    Exit Sub
  End If

  VERIFICAJAPAGO(CanContinue)
  If CanContinue = False Then
    BsShowMessage("Alteração não permitida nesta fase","E")
  End If
End Sub

Public Sub TABLE_AfterEdit()

  If SessionVar("Reconsiderar") = "S" Then
    ReconsiderarGlosa
  Else
    CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
    CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
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

Public Function VerificaAgrupadorPagamentoFechado As Boolean
	Dim callEntity As CSEntityCall
  	Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamGuiaEventosGlosa, Benner.Saude.Entidades", "VerificaAgrupadorPagamentoFechado")
  	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger)
  	VerificaAgrupadorPagamentoFechado = CBool(callEntity.Execute)
	Set callEntity =  Nothing
End Function

Public Sub PreencheSessionVar
  Dim Peg As Object

  Set Peg = NewQuery
  Peg.Add("SELECT G.PEG")
  Peg.Add("  FROM SAM_GUIA_EVENTOS_GLOSA EGL					   ")
  Peg.Add("  JOIN SAM_GUIA_EVENTOS EG ON EG.HANDLE = EGL.GUIAEVENTO")
  Peg.Add("  JOIN SAM_GUIA G ON G.HANDLE = EG.GUIA				   ")
  Peg.Add(" WHERE EGL.HANDLE = :HANDLE                             ")
  Peg.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Peg.Active = True

  SessionVar("HPEG") = Peg.FieldByName("PEG").AsString

  Set Peg = Nothing
End Sub
