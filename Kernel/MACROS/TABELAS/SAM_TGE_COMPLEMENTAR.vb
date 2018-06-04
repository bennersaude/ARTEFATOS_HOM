'HASH: 98B60F2D47089A769B19AFE9C4919131
'Macro: SAM_TGE_COMPLEMENTAR

'#Uses "*bsShowMessage"

Public Sub EVENTOAGERAR_OnExit()
  Dim SQL As BPesquisa
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT CIRURGICO FROM SAM_TGE WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  SQL.Active = True
  If SQL.FieldByName("CIRURGICO").AsString = "S" Then
    CODIGOPAGTO.ReadOnly = True
  Else
    CODIGOPAGTO.ReadOnly = False
  End If
  Set SQL = Nothing
End Sub

Public Sub GRAUAGERAR_OnChange()
  Dim Q As BPesquisa
  Set Q = NewQuery
  Q.Add("SELECT COUNT (*) REC FROM SAM_TGE_GRAU WHERE EVENTO = :EVENTO AND GRAU = :GRAU")
  Q.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  Q.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAUAGERAR").AsInteger
  Q.Active = True
  If Q.FieldByName("REC").AsInteger = 0 Then
    CurrentQuery.FieldByName("GRAUAGERAR").Clear
  End If

  Set Q = Nothing
End Sub

Public Sub GRAUAGERAR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim qParamAtend As BPesquisa

  Set qParamAtend = NewQuery

  qParamAtend.Add("SELECT FILTRARGRAUSVALIDOS FROM SAM_PARAMETROSATENDIMENTO")
  qParamAtend.Active = True

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "GRAU|DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  If CurrentQuery.FieldByName("EVENTOAGERAR").IsNull Then
    vCriterio = "HANDLE = -1"
  Else
  	If qParamAtend.FieldByName("FILTRARGRAUSVALIDOS").AsString = "S" Then
      vCriterio = "HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + ")"
  	Else
  	  vCriterio = ""
  	End If
  End If

  vCampos = "Grau|Descrição|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, vCriterio, "Tabela De Graus", False, GRAUAGERAR.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAUAGERAR").Value = vHandle
  End If
  Set interface = Nothing

  Set qParamAtend = Nothing
End Sub

Public Sub TABLE_AfterScroll()

  Dim qParamAtend As BPesquisa
  Set qParamAtend = NewQuery

  If (WebMode) Then
    qParamAtend.Add("SELECT FILTRARGRAUSVALIDOS FROM SAM_PARAMETROSATENDIMENTO")
    qParamAtend.Active = True

    If qParamAtend.FieldByName("FILTRARGRAUSVALIDOS").AsString = "S" Then
      GRAUAGERAR.WebLocalWhere = "A.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTOAGERAR))"
    Else
      GRAUAGERAR.WebLocalWhere = ""
    End If

    EVENTOAGERAR.WebLocalWhere = "A.ULTIMONIVEL = 'S' "
  End If

  Set qParamAtend = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qParam As BPesquisa
  Set qParam = NewQuery

  qParam.Active = False
  qParam.Clear
  qParam.Add("SELECT CALCCODPAGTOEVENTOCIRURGICO, FORNECIMENTOMEDICAMENTO FROM SAM_PARAMETROSATENDIMENTO")
  qParam.Active = True

  Dim qEvento As BPesquisa
  Set qEvento = NewQuery

  qEvento.Clear
  qEvento.Add("SELECT TABTIPOEVENTO FROM SAM_TGE WHERE HANDLE = :HEVENTO")

  If WebMode Or VisibleMode Then
  	qEvento.ParamByName("HEVENTO").Value = RecordHandleOfTable("SAM_TGE")
  Else
  	qEvento.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
  End If
  qEvento.Active = True


  If qParam.FieldByName("FORNECIMENTOMEDICAMENTO").AsString <> "N" And _
     qEvento.FieldByName("TABTIPOEVENTO").AsInteger = 4 Then
     If VisibleMode Or WebMode Then
       If RecordHandleOfTable("SAM_TGE") <> CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger Then
	     CanContinue = False
	     Set qEvento = Nothing
	     bsShowMessage("Evento do tipo 'Medicamento' não possibilita 'Evento Complementar' diferente dele mesmo!","E")
	     Exit Sub
       End If
     Else
       If CurrentQuery.FieldByName("EVENTO").AsInteger <> CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger Then
    	 CanContinue = False
    	 Set qEvento = Nothing
    	 bsShowMessage("Evento do tipo 'Medicamento' não possibilita 'Evento Complementar' diferente dele mesmo!","E")
    	 Exit Sub
       End If
     End If
  End If



  Set qEvento = Nothing

  Dim SQL As BPesquisa
  Set SQL = NewQuery

  SQL.Add("SELECT COUNT(*) T                    ")
  SQL.Add("  FROM SAM_TGE_COMPLEMENTAR          ")
  SQL.Add(" WHERE EVENTOAGERAR = :EVENTOAGERAR  ")
  SQL.Add("   AND GRAUAGERAR = :GRAUAGERAR      ")
  SQL.Add("   AND EVENTO = :EVENTO              ")

  If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    SQL.Add("   AND HANDLE <> :HEVENTOCOMPLEMENTAR")
    SQL.ParamByName("HEVENTOCOMPLEMENTAR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If

  SQL.ParamByName("EVENTOAGERAR").AsInteger = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  SQL.ParamByName("GRAUAGERAR").AsInteger   = CurrentQuery.FieldByName("GRAUAGERAR").AsInteger
  SQL.ParamByName("EVENTO").AsInteger       = CurrentQuery.FieldByName("EVENTO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("T").AsInteger > 0 Then
      CanContinue = False
      Set SQL = Nothing
      bsShowMessage("Registro Duplicado! Operação não permitida.","E")
      Exit Sub
  End If
  Set SQL = Nothing


  Dim qEventoGerar As BPesquisa
  Set qEventoGerar = NewQuery

  qEventoGerar.Clear
  qEventoGerar.Add("SELECT TABTIPOEVENTO, CIRURGICO FROM SAM_TGE WHERE HANDLE = :HEVENTO")
  qEventoGerar.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger
  qEventoGerar.Active = True

  'Se o evento é cirúrgico não pode ser informado o codigo do pagamento
  If (qParam.FieldByName("CALCCODPAGTOEVENTOCIRURGICO").AsString = "S" And _
      qEventoGerar.FieldByName("CIRURGICO").AsString = "S" And _
      Not(CurrentQuery.FieldByName("CODIGOPAGTO").IsNull)) Then

      Dim msg As String
      CanContinue = False
      msg = "Está marcado nos parâmetros gerais que o percentual de pagamento" + Chr(13)
      msg = msg + "será calculado pelo sistema para eventos cirúrgicos." + Chr(13)
      msg = msg + "O campo Código de pagamento deverá ser deixado em branco!"

      Set qEventoGerar = Nothing
      Set qParam = Nothing
      bsShowMessage(msg, "E")
      Exit Sub
  End If

  	If (qParam.FieldByName("FORNECIMENTOMEDICAMENTO").AsString <> "N" And _
        qEventoGerar.FieldByName("TABTIPOEVENTO").AsInteger = 4) Then
      If VisibleMode Or WebMode Then
        If RecordHandleOfTable("SAM_TGE") <> CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger Then
		  CanContinue = False
      	  Set qEventoGerar = Nothing
      	  Set qParam = Nothing
      	  bsShowMessage("Evento do tipo 'Medicamento' não pode ser 'Evento Complementar'!", "E")
      	  Exit Sub
        End If
      Else
        If CurrentQuery.FieldByName("EVENTO").AsInteger <> CurrentQuery.FieldByName("EVENTOAGERAR").AsInteger Then
		  CanContinue = False
      	  Set qEventoGerar = Nothing
          Set qParam = Nothing
       	  bsShowMessage("Evento do tipo 'Medicamento' não pode ser 'Evento Complementar'!", "E")
      	  Exit Sub
      	End If
	  End If
    End If

  Set qEventoGerar = Nothing
  Set qParam = Nothing

  If CurrentQuery.FieldByName("QTD").AsInteger > 1 Then
    Dim qGrau As BPesquisa
    Set qGrau = NewQuery

    qGrau.Clear
    qGrau.Add("SELECT ORIGEMVALOR FROM SAM_GRAU WHERE HANDLE = :GRAU")
    qGrau.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAUAGERAR").AsInteger
    qGrau.Active = True

    If qGrau.FieldByName("ORIGEMVALOR").AsInteger = 2 Then
      CanContinue = False
      Set qGrau = Nothing
      bsShowMessage("Somente 1 auxiliar de cada tipo é permitido.", "E")
      Exit Sub
    End If

    Set qGrau = Nothing
  End If

End Sub

Public Sub EVENTOAGERAR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO|NIVELAUTORIZACAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição|Nível"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", False, EVENTOAGERAR.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOAGERAR").Value = vHandle
    CurrentQuery.FieldByName("GRAUAGERAR").Value = Null
  End If
  Set interface = Nothing

End Sub


Public Sub CODIGOPAGTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PERCENTUALPGTO.CODIGOPAGTO|DESCRICAO|INCIDENCIAMINIMA|PERCENTUALPGTOINCIDENCIA1|PERCENTUALPGTODEMAIS|USADOAUTORIZACAO|USADOPAGAMENTO"

  vCampos = "Código|Descrição|Incidência Mínima|% Pagto Inc 1|% Pagto Demais|Usado Autorização|Usado Pagto"

  vHandle = interface.Exec(CurrentSystem, "SAM_PERCENTUALPGTO", vColunas, 1, vCampos, vCriterio, "Tabela de Códigos de Pagamentos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CODIGOPAGTO").Value = vHandle
  End If

  Set interface = Nothing


End Sub




