'HASH: 1A2D59775581E82D60D34B47AF9C4836
'#Uses "*bsShowMessage"

Public Sub Main

	UserVar("Sender") = "BOTAOEXPORTAR"

    SessionVar("VIAAGENDAMENTO") = "S"

	Dim diretorioExportacaoCartao As String
    diretorioExportacaoCartao = SessionVar("DIRETORIOEXPORTACAOCARTAO")

    If diretorioExportacaoCartao = "" Then
      BsShowMessage("Diretório para exportação dos arquivos não foi parametrizado [DIRETORIOEXPORTACAOCARTAO].", "E")
      Exit Sub
    End If

    If Len(diretorioExportacaoCartao) > 100 Then
      BsShowMessage("Diretório para exportação dos arquivos junto do nome do arquivo não pode ultrapassar 100 caracteres.", "E")
      Exit Sub
    End If

	Dim qRotinas As BPesquisa
	Set qRotinas = NewQuery

    qRotinas.Active = False
    qRotinas.Add("SELECT *                      ")
    qRotinas.Add("  FROM SAM_ROTINACARTAO       ")
    qRotinas.Add(" WHERE SITUACAOEXPORTACAO = 1 ")
    qRotinas.Add("   AND SITUACAO = '5'         ")
    qRotinas.Add("   AND TABTIPOROTINA = 1      ")
    qRotinas.Add("   AND TABTIPOLEIAUTE = 2     ")
    qRotinas.Active = True

    Dim ParametrosBenef As BPesquisa

	Set ParametrosBenef = NewQuery

	ParametrosBenef.Active = False
	ParametrosBenef.Add(" SELECT * 							")
	ParametrosBenef.Add("   FROM SAM_PARAMETROSBENEFICIARIO ")
	ParametrosBenef.Active = True

	Set Obj = CreateBennerObject("BSInterface0004.RotinaCartao")

	While(Not(qRotinas.EOF))
  		UserParam = qRotinas.FieldByName("HANDLE").AsInteger

    	Obj.Exportar(CurrentSystem, qRotinas.FieldByName("HANDLE").AsInteger)

		If (ParametrosBenef.FieldByName("DESBLOQUEARAPOSEXPORTACAO").AsString = "S") Then

    		Obj.Desbloquear(CurrentSystem, qRotinas.FieldByName("HANDLE").AsInteger)

		End If

		qRotinas.Next

	Wend


	Set ParametrosBenef = Nothing

	SessionVar("VIAAGENDAMENTO") = "N"

	Set qRotinas = Nothing
	Set Obj = Nothing
End Sub
