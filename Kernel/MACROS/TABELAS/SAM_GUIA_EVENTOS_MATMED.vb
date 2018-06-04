'HASH: 81AEBA8CFEA7593EED608CCF122E4699

'#Uses "*bsShowMessage"
'#Uses "*VerificaPermissaoEdicaoTriagem"
'#Uses "*LimpaEspaco"

Option Explicit

Dim gGuiaEvento As Long


Public Sub reprocessa(pGuiaEvento As Long)
  'SE ESTIVER EM DIGITACAO NAO FAZ NADA
  Dim sql As Object
  Set sql = NewQuery
  sql.Add("SELECT G.SITUACAO FROM SAM_GUIA G WHERE HANDLE=" + Str(RecordHandleOfTable("SAM_GUIA")))
  sql.Active = True

  If sql.FieldByName("SITUACAO").AsString <>"1" Then
    Dim Interface As Object
    Set Interface = CreateBennerObject("SAMPEG.Rotinas")
    Interface.RevisarEvento(CurrentSystem, pGuiaEvento, "PARCIAL", True)
    Set Interface = Nothing
    If VisibleMode Then
    	CurrentQuery.Active = False
    	CurrentQuery.Active = True
    End If

 End If
  Set sql = Nothing
End Sub



Public Sub BOTAORECONSIDERAR_OnClick()
  'SMS 63895 - Marcelo Barbosa - 27/06/2006
  Dim sql As Object
  Set sql = NewQuery
  sql.Add("SELECT G.SITUACAO FROM SAM_GUIA G WHERE G.HANDLE=" + Str(RecordHandleOfTable("SAM_GUIA")))
  sql.Active = True

  If sql.FieldByName("SITUACAO").AsString <> "1" Then
    QTDRECONSIDERADA.ReadOnly = False
    VALORRECONSIDERADO.ReadOnly = False
    QTDRECONSIDERADA.SetFocus
    CurrentQuery.Edit
  End If
End Sub

Public Sub QTDRECONSIDERADA_OnExit()
  If CurrentQuery.State <>1 Then
    CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = (CurrentQuery.FieldByName("VALORINFORMADO").AsFloat / CurrentQuery.FieldByName("QTDINFORMADA").AsFloat) * CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat
  End If
End Sub

Public Sub TABLE_AfterDelete()
  reprocessa(gGuiaEvento)
End Sub

Public Sub TABLE_AfterEdit()
	If SessionVar("Reconsiderar") = "S" Then
		SessionVar("Reconsiderar") = "N"
	End If
End Sub

Public Sub TABLE_AfterPost()
  Dim sql As Object
  Set sql = NewQuery
  sql.Add("SELECT QTDAPRESENTADA FROM SAM_GUIA_EVENTOS WHERE HANDLE=:HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
  sql.Active = True

  If sql.FieldByName("QTDAPRESENTADA").AsFloat <>1 Then
    bsShowMessage("A quantidade do evento será automaticamente alterada para 1", "I")
    sql.Clear
    sql.Add("UPDATE SAM_GUIA_EVENTOS SET QTDAPRESENTADA=1 WHERE HANDLE=:HANDLE")
    sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
    sql.ExecSQL
  End If

  Set sql = Nothing
  reprocessa(CurrentQuery.FieldByName("GUIAEVENTO").AsInteger)
End Sub

Public Sub TABLE_AfterScroll()
  gGuiaEvento = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
  QTDRECONSIDERADA.ReadOnly = True
  VALORRECONSIDERADO.ReadOnly = True
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If Not VerificarPermissaoUsuarioPegTriado(True) Then
 	CanContinue = False
 	Exit Sub
  End If

  'SMS 68864 - Débora Rebello - 05/03/2007
  If isTipoPegReembolsoMatMed(CurrentQuery.FieldByName("GUIAEVENTO").AsFloat) Then
    bsShowMessage("Não é permitido excluir Materiais/Medicamentos de PEGs do tipo Reembolso de Mat/Med", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If Not VerificarPermissaoUsuarioPegTriado(True) Then
 	  CanContinue = False
 	  Exit Sub
    End If

	If SessionVar("Reconsiderar") = "S" Then 'SMS 90442 - Ricardo Rocha - 25/04/2008
		'SMS 68864 - Débora Rebello - 05/03/2007
		If isTipoPegReembolsoMatMed(CurrentQuery.FieldByName("GUIAEVENTO").AsFloat) Then
			bsShowMessage("Não é permitido alterar Materiais/Medicamentos de PEGs do tipo Reembolso de Mat/Med", "E")
			CanContinue = False
			Exit Sub
		End If
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim sql As Object

  If Not VerificarPermissaoUsuarioPegTriado(True) Then
 	CanContinue = False
 	Exit Sub
  End If

  'SMS 68864 - Débora Rebello - 05/03/2007 - inicio
  If isTipoPegReembolsoMatMed(RecordHandleOfTable("SAM_GUIA_EVENTOS")) Then
    bsShowMessage("Não é permitido inserir Materiais/Medicamentos em eventos de PEGs do tipo Reembolso de Mat/Med", "E")
    CanContinue = False
    Exit Sub
  End If
  'SMS 68864 - Débora Rebello - 05/03/2007 - fim

  Set sql = NewQuery
  sql.Add("SELECT T.CLASSIFICACAO FROM SAM_GUIA_EVENTOS E, SAM_GRAU G, SAM_TIPOGRAU T WHERE E.HANDLE=:H AND G.HANDLE=E.GRAU AND T.HANDLE=G.TIPOGRAU")
  sql.ParamByName("H").AsInteger = RecordHandleOfTable("SAM_GUIA_EVENTOS")
  sql.Active = True
  If sql.FieldByName("CLASSIFICACAO").AsString = "1" Or sql.FieldByName("CLASSIFICACAO").AsString = "2" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage("O grau do evento deve ser do tipo Material ou Medicamento", "E")
    Exit Sub
  End If
  Set sql = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim vLstMatMed As Object
	Dim qAux As Object
	Dim qAux1 As Object
	Dim vdValor As Double
	Dim vsDescricao As String
	Dim vsMsg As String
	Dim viResult As Long

    If Not CurrentQuery.FieldByName("NOTAFISCAL").IsNull Then
      CurrentQuery.FieldByName("NOTAFISCAL").AsString = LimpaEspaco(CurrentQuery.FieldByName("NOTAFISCAL").AsString)
    End If

	'SMS 68864 - Débora Rebello - 05/03/2007
	If isTipoPegReembolsoMatMed(CurrentQuery.FieldByName("GUIAEVENTO").AsFloat) Then
		bsShowMessage("Não é permitido alterar Materiais/Medicamentos de PEGs do tipo Reembolso de Mat/Med", "E")
		CanContinue = False
		Exit Sub
	End If

	If CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat >(CurrentQuery.FieldByName("QTDGLOSADA").AsFloat)Then
		bsShowMessage("A quantidade reconsiderada não pode ser maior que a quantidade glosada", "E")
		CanContinue = False
		Exit Sub
	End If

	If CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat >(CurrentQuery.FieldByName("VALORGLOSADO").AsFloat)Then
		bsShowMessage("O valor reconsiderado não pode ser maior que o valor glosado", "E")
		CanContinue = False
		Exit Sub
	End If

	'Verificar se a hora está perto da meia noite
	'Para evitar que o problema de um horário que comece num dia e termine em outro
	'foi criada uma janela que permite a inclusão de horários a partir de 20h até as 0h

	If CurrentQuery.FieldByName("HORAINICIAL").AsDateTime >= TimeValue("20:00") And _
	   CurrentQuery.FieldByName("HORAINICIAL").AsDateTime <= TimeValue("00:00") Then
		'Passou de 20:00 pode ser que o horário seguinte seja depois da meia noite
		If Not ((CurrentQuery.FieldByName("HORAFINAL").AsDateTime >= TimeValue("00:00")) And _
		   (CurrentQuery.FieldByName("HORAFINAL").AsDateTime <= TimeValue("06:00"))) Then
			bsShowMessage("A hora inicial não pode ser maior que a hora final.", "E")
			CanContinue = False
			Exit Sub
		End If
	ElseIf CurrentQuery.FieldByName("HORAINICIAL").AsDateTime > CurrentQuery.FieldByName("HORAFINAL").AsDateTime Then
		bsShowMessage("A hora inicial não pode ser maior que a hora final.", "E")
		CanContinue = False
		Exit Sub
	End If

	Set qAux = NewQuery
  	Set qAux1 = NewQuery

	qAux.Clear
	qAux.Add("SELECT SALBASE FROM SAM_MATMED WHERE HANDLE = :HANDLE")
	qAux.Active = False
	qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MATMED").AsInteger
	qAux.Active = True

	qAux1.Clear
	qAux1.Add("SELECT G.PEG, GE.HANDLE")
	qAux1.Add("  FROM SAM_GUIA G, SAM_GUIA_EVENTOS GE")
	qAux1.Add(" WHERE G.HANDLE = GE.GUIA")
	qAux1.Add("   AND GE.HANDLE = :HANDLE")
	qAux1.Active = False
	qAux1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
	qAux1.Active = True

	If qAux.FieldByName("SALBASE").AsInteger > 0 Then
		Set vLstMatMed = CreateBennerObject("BSPRO006.Geral")
		viResult = vLstMatMed.ListarSalMed(CurrentSystem, qAux.FieldByName("SALBASE").AsInteger, qAux1.FieldByName("HANDLE").AsInteger, vdValor, vsDescricao)

		If (viResult > 0) And (viResult <> CurrentQuery.FieldByName("MATMED").AsInteger) Then
			vsMsg = "Existe outro material/medicamento com o mesmo Sal Base e com valor menor" + Chr(13) + Chr(13) + _
					"Mat/Med: " + vsDescricao + Chr(13) + "Valor: " + CStr(vdValor)

			If CurrentQuery.FieldByName("JUSTIFICATIVA").IsNull Then
				vsMsg = vsMsg + Chr(13) + Chr(13) + "Necessário informar uma justificativa"
				bsShowMessage(vsMsg, "E")
				CanContinue = False
				Exit Sub
			End If
		End If
	End If

End Sub

Public Function isTipoPegReembolsoMatMed(GuiaEventos As Double) As Boolean
  'SMS 68864 - Débora Rebello - 05/03/2007
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT TP.REEMBOLSOMATMED ")
  sql.Add("  FROM SAM_TIPOPEG TP ")
  sql.Add("  JOIN SAM_PEG P ON P.TIPOPEG = TP.HANDLE ")
  sql.Add("  JOIN SAM_GUIA G ON G.PEG = P.HANDLE ")
  sql.Add("  JOIN SAM_GUIA_EVENTOS GE ON GE.GUIA = G.HANDLE ")
  sql.Add(" WHERE GE.HANDLE = :HANDLE ")
  sql.ParamByName("HANDLE").AsInteger = GuiaEventos
  sql.Active = True

  isTipoPegReembolsoMatMed = (sql.FieldByName("REEMBOLSOMATMED").AsString = "S")

  Set sql = Nothing
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
'SMS 90442 - Ricardo Rocha - 25/04/2008
	If CommandID = "BOTAORECONSIDERAR" Then
		SessionVar("Reconsiderar") = "S"
	End If
End Sub
