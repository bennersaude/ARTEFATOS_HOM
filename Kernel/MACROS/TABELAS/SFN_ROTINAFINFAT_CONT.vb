'HASH: FC0BEE6641A9A6D85225C9430F7D9F2F
'Macro: SFN_ROTINAFINFAT_CONT
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0016.RotinaFaturamentoBeneficiarios")
    Obj.CancelarFaturamento(CurrentSystem, RecordHandleOfTable("SFN_ROTINAFINFAT"), "C", CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
  	If bsShowMessage("Cancelar Faturamento?", "Q") = vbYes Then
	    Dim vsMensagemErro As String
    	Dim viRetorno As Long
    	Dim vcContainer As CSDContainer

	    Set vcContainer = NewContainer

	    vcContainer.AddFields("HANDLE:INTEGER")
    	vcContainer.AddFields("OPCAOCANCELAMENTO:STRING")
    	vcContainer.AddFields("HOPCAOCANCELAMENTO:INTEGER")

	    vcContainer.Insert
    	vcContainer.Field("HANDLE").AsInteger             = RecordHandleOfTable("SFN_ROTINAFINFAT")
    	vcContainer.Field("OPCAOCANCELAMENTO").AsString   = "C"
    	vcContainer.Field("HOPCAOCANCELAMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
    									"BSBen018", _
    									"RotinaFaturamentoBeneficiarios_Cancelar", _
    									"Cancelando contrato " + CurrentQuery.FieldByName("CONTRATO").AsString, _
    									0, _
    									"", _
    									"", _
    									"", _
    									"", _
    									"C", _
    									False, _
    									vsMensagemErro, _
    									vcContainer)
    	If viRetorno = 0 Then
      	bsShowMessage("Processo enviado para execução no servidor!", "I")
    	Else
      	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	End If
    End If

  End If
  Set Obj = Nothing
End Sub

Public Sub BOTAORH_OnClick()
  Dim Obj As Object
  Dim SQL As Object
  Dim SQL2 As Object
  Dim vCompetencia As Date
  Dim vDataRotina As Date
  Dim vTipoFaturamento As Integer
  Dim vRotinaFinFat As Integer
  Dim VRotFinAux As Integer
  Dim VContratoAux As Integer

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  VContratoAux = CurrentQuery.FieldByName("CONTRATO").Value

  Set SQL2 = NewQuery

  SQL2.Add("SELECT ROTINAFIN FROM SFN_ROTINAFINFAT WHERE HANDLE =:ROTFAT")

  SQL2.ParamByName("ROTFAT").Value = CurrentQuery.FieldByName("ROTINAFINFAT").Value
  SQL2.Active = True

  VRotFinAux = SQL2.FieldByName("ROTINAFIN").AsInteger

  Set SQL = NewQuery

  SQL.Add("SELECT A.COMPETENCIA COMPETENCIA, B.DESCRICAO, B.DATAROTINA DATAROTINA, B.SEQUENCIA SEQROTINA, D.HANDLE TIPOFATURAMENTO")
  SQL.Add("FROM SFN_COMPETFIN A, SFN_ROTINAFIN B, SIS_TIPOFATURAMENTO D")
  SQL.Add("WHERE B.HANDLE =:ROTFINFAT AND B.COMPETFIN = A.HANDLE AND A.TIPOFATURAMENTO = D.HANDLE")

  SQL.ParamByName("ROTFINFAT").Value = VRotFinAux
  SQL.Active = True

  vCompetencia = SQL.FieldByName("COMPETENCIA").AsDateTime
  vDataRotina = SQL.FieldByName("DATAROTINA").AsDateTime
  vTipoFaturamento = SQL.FieldByName("TIPOFATURAMENTO").AsInteger
  vRotinaFinFat = CurrentQuery.FieldByName("ROTINAFINFAT").Value

  If VisibleMode Then
    Set Obj = CreateBennerObject("RotArq.Rotinas")
    Obj.ArquivoRh(CurrentSystem, vCompetencia, vDataRotina, vTipoFaturamento, vRotinaFinFat, VContratoAux)
    Set Obj = Nothing
  Else
    Dim vsMensagemErro As String
    Dim viRetorno As Long
    Dim vcContainer As CSDContainer

    Set vcContainer = NewContainer
    vcContainer.AddFields("COMPETENCIA:TDATETIME;DATAROTINA:TDATETIME;TIPOFATURAMENTO:INTEGER;ROTINAFINFAT:INTEGER;CONTRATO:INTEGER")
    vcContainer.Insert
    vcContainer.Field("COMPETENCIA").AsDateTime = vCompetencia
    vcContainer.Field("DATAROTINA").AsDateTime = vDataRotina
    vcContainer.Field("TIPOFATURAMENTO").AsInteger = vTipoFaturamento
    vcContainer.Field("ROTINAFINFAT").AsInteger = vRotinaFinFat
    vcContainer.Field("CONTRATO").AsInteger = VContratoAux

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                                "RotArq", _
	                                "ArquivoRH_Exec", _
	                                "Geração de Arquivo para RH: " + _
	                                SQL.FieldByName("COMPETENCIA").AsString + " - " + _
									SQL.FieldByName("SEQROTINA").AsString + " - " + _
									SQL.FieldByName("DESCRICAO").AsString, _
	                                0, _
	                                "", _
	                                "", _
	                                "", _
	                                "", _
	                                "P", _
	                                False, _
	                                vsMensagemErro, _
	                                vcContainer)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set Obj = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAORH"
			BOTAORH_OnClick
	End Select
End Sub
