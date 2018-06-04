﻿'HASH: 351F7F127B59CEA6243F37876724A563
'TV_FORM0092
'#Uses "*bsShowMessage"
'#Uses "*RecordHandleOfTableInterfacePEG"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  On Error GoTo Erro:
	  If CurrentQuery.FieldByName("VALORINFORMADOPF").AsFloat < 0 Then
	    BsShowMessage("Valor informado deve ser maior ou igual a 0", "E")
	    CanContinue = False
	    CurrentQuery.FieldByName("VALORINFORMADOPF").AsFloat = 0
	    Exit Sub
	  End If

	  Dim handleGuiaEvento As Long
	  Dim qGuiaEventos As BPesquisa
	  Dim qParamGerais As BPesquisa
	  Dim qUpdateEvento As BPesquisa
	  Set qGuiaEventos = NewQuery
	  Set qParamGerais = NewQuery
	  Set qUpdateEvento = NewQuery

	  handleGuiaEvento = RecordHandleOfTableInterfacePEG("SAM_GUIA_EVENTOS")

	  qGuiaEventos.Clear
	  qGuiaEventos.Active = False
	  qGuiaEventos.Add("  SELECT VALORPROVISAOPAGAMENTO  ")
	  qGuiaEventos.Add("    FROM SAM_GUIA_EVENTOS        ")
	  qGuiaEventos.Add("   WHERE HANDLE = :HANDLE        ")
	  qGuiaEventos.ParamByName("HANDLE").AsInteger = handleGuiaEvento
	  qGuiaEventos.Active = True

	  qParamGerais.Clear
	  qParamGerais.Active = False
	  qParamGerais.Add("  SELECT * FROM SAM_PARAMETROSPROCCONTAS  ")
	  qParamGerais.Active = True

	  If CurrentQuery.FieldByName("VALORINFORMADOPF").AsFloat > (qGuiaEventos.FieldByName("VALORPROVISAOPAGAMENTO").AsFloat * (qParamGerais.FieldByName("PFMAXIMA").AsFloat/100)) Then
	    bsShowMessage("Valor PF ultrapassou o máximo permitido.", "E")
	    CanContinue = False
	    Set qGuiaEventos = Nothing
	    Set qParamGerais = Nothing
	    Exit Sub
	  End If

	  qUpdateEvento.Clear
	  qUpdateEvento.Active = False
	  qUpdateEvento.Add("  UPDATE SAM_GUIA_EVENTOS           ")
	  qUpdateEvento.Add("     SET VALORINFORMADOPF = :VALOR  ")
	  qUpdateEvento.Add("   WHERE HANDLE = :HANDLE           ")
	  qUpdateEvento.ParamByName("VALOR").AsFloat = CurrentQuery.FieldByName("VALORINFORMADOPF").AsFloat
	  qUpdateEvento.ParamByName("HANDLE").AsInteger = handleGuiaEvento
	  qUpdateEvento.ExecSQL

	  Dim interface As Object
	  Set interface = CreateBennerObject("SAMPEG.ROTINAS")
	  interface.RevisarEvento(CurrentSystem, handleGuiaEvento, "TOTAL", True)

	  bsShowMessage("Valor PF alterado com sucesso", "I")

	  Set interface = Nothing
	  Set qGuiaEventos = Nothing
	  Set qParamGerais = Nothing
	  Set qUpdateEvento = Nothing
	  Exit Sub

  Erro:
    bsShowMessage("Erro: Não foi possível alterar Valor PF", "E")
    Set interface = Nothing
    Set qGuiaEventos = Nothing
    Set qParamGerais = Nothing
    Set qUpdateEvento = Nothing


End Sub
