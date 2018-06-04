'HASH: 4FB6D3A2166C282A5124D8DA9B97F109
'#Uses "*bsShowMessage"


Public Sub BOTAOMOVFIN_OnClick()
	Dim qAux As Object

	Set qAux = NewQuery

	qAux.Add("SELECT HANDLE FROM R_RELATORIOS")
	qAux.Add(" WHERE CODIGO = 'BEN070'")
 	qAux.Active = True

	ReportPreview(qAux.FieldByName("HANDLE").AsInteger, "", False, False)

 	Set qAux = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	Dim Nulo As String

	If (StrPos("ORACLE", SQLServer) > 0) Then
		Nulo = "NVL"
	ElseIf (StrPos("DB2", SQLServer) > 0) Then
		Nulo = "COALESCE"
	Else
		Nulo = "ISNULL"
	End If


	Dim vCriterio As String
	vCriterio = "HANDLE NOT IN (SELECT " + Nulo + "(MOTIVOFALECIMENTOTITULAR,0) FROM SAM_PARAMETROSBENEFICIARIO " + _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOFALECIMENTO,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOMIGRACAO,0) FROM SAM_PARAMETROSBENEFICIARIO " + _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MIGRACAONREGREG,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOMIGRACAOCORRESPONSAVEL,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(ADAPTACAO,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOCANCELAMENTOINDICADOR,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOFALTADEPENDENTEAGREGADO,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOCANCMIGRATIVOINATIVO,0) FROM SAM_PARAMETROSBENEFICIARIO)" _

	If VisibleMode Then
		MOTIVO.LocalWhere = vCriterio
	ElseIf WebMode Then
		MOTIVO.WebLocalWhere = vCriterio
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If VisibleMode Then
    If bsShowMessage("Deseja realmente cancelar?","Q") = vbNo Then
      CanContinue = False
      Exit Sub
    End If
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT DATACANCELAMENTO")
  SQL.Add("  FROM SAM_FAMILIA    ")
  SQL.Add(" WHERE HANDLE = :HANDLE")

  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_FAMILIA")

  SQL.Active = True

  Dim viTitular As Long
  Dim vbCancelar As Boolean
  Dim Obj As Object
  Set Obj = CreateBennerObject("SAMCANCELAMENTO.Cancelar")

  viTitular = Obj.BuscaTitular(CurrentSystem, RecordHandleOfTable("SAM_FAMILIA"))

  If Obj.ChecaInternado(CurrentSystem, viTitular, CurrentQuery.FieldByName("DATA").AsDateTime) = True Then
	If bsShowMessage("Beneficiário titular com autorização de internação em aberto. Deseja realmente cancelar?", "Q") = vbYes Then
	  vbCancelar = True
	Else
	  vbCancelar = False
	  bsShowMessage("Beneficiário titular com autorização de internação em aberto. Cancelamento abortado.", "E")
	End If
  Else
	vbCancelar = True
  End If

  If vbCancelar = True Then
  	If VisibleMode Then
	  bsShowMessage(Obj.Familia(CurrentSystem, CLng(SessionVar("HFAMILIA_CANCELAMENTO")),CurrentQuery.FieldByName("DATA").AsDateTime,CurrentQuery.FieldByName("CHECKCANCFATDOC").AsBoolean ,CurrentQuery.FieldByName("MOTIVO").AsInteger, CurrentQuery.FieldByName("DATAATEND").AsDateTime), "I")
  	ElseIf WebMode Then
  	  If SQL.FieldByName("DATACANCELAMENTO").IsNull Then
		bsShowMessage(Obj.Familia(CurrentSystem, RecordHandleOfTable("SAM_FAMILIA") ,CurrentQuery.FieldByName("DATA").AsDateTime,CurrentQuery.FieldByName("CHECKCANCFATDOC").AsBoolean, CurrentQuery.FieldByName("MOTIVO").AsInteger, CurrentQuery.FieldByName("DATAATEND").AsDateTime),  "I")
 	  Else
	 	bsShowMessage("Familia já cancelada!","I")
 	  End If
    End If
  End If

  Set Obj = Nothing



End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOMOVFIN" Then
		BOTAOMOVFIN_OnClick
	End If
End Sub
