'HASH: 38885BA6620E0EE935A191A6C857239C
'Macro: ANS_SIP_ANEXO
'#Uses "*bsShowMessage"


Public Sub ANEXO_OnChange()
CurrentQuery.UpdateRecord
If CurrentQuery.FieldByName("ANEXO").AsString = "1" Then
    RELATORIOANEXOI.Visible = True
  Else
    RELATORIOANEXOI.Visible = False
  End If
End Sub

Public Sub BOTAOCRIARITENS_OnClick()
  Dim Obj As Object
  Dim vsMensagemRetorno As String
  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If
  Set Obj = CreateBennerObject("BSANS001.CriarItens")
  Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger,vsMensagemRetorno)
  Set Obj = Nothing

  If Not vsMensagemRetorno = "" Then
	bsShowMessage(vsMensagemRetorno,"I")
  End If

End Sub

Public Sub BOTAODUPLICARANEXO_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If
  Set Obj = CreateBennerObject("BSANS001.uRotinas")
  Obj.DuplicarModelo(CurrentSystem)
  Set Obj = Nothing

End Sub

Public Sub BOTAOEXCLUIRITENS_OnClick()
	If bsShowMessage("Deseja excluir todos os itens? " + (Chr(13)) + " Continuar?", "Q") = vbYes Then
		Dim query As Object
		Set query = NewQuery

		query.Active = False
		query.Clear
		query.Add("SELECT CASE                                                   ")
		query.Add("          WHEN COUNT(1) > 0 THEN 'S'                          ")
		query.Add("          ELSE 'N'                                            ")
		query.Add("       END AS USOU                                            ")
		query.Add("  FROM ANS_SIP_COMPETENCIA A                                  ")
		query.Add("  JOIN ANS_SIP_COMPETENCIA_ITEM B ON (B.SIPCOMPET = A.HANDLE) ")
		query.Add(" WHERE ANEXO = :HANDLE")
		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
		query.Active = True

		If query.FieldByName("USOU").AsString = "S" Then
			 bsShowMessage("Existem anexos gerados com os itens deste parametro.", "E")
			 Exit Sub
		End If

		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_EVENTO                               			")
		query.Add(" WHERE SIPITEM IN (                                                 			")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

  		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_CARENCIA                            			")
		query.Add(" WHERE SIPANEXO IN (                                                 		")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

  		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_MODULO                               			")
		query.Add(" WHERE SIPANEXO IN (                                                 		")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

  		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_REGIMEATEND                               	")
		query.Add(" WHERE SIPITEM IN (                                                 			")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

  		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_TIPOTRAT                               		")
		query.Add(" WHERE SIPANEXO IN (                                                			")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

  		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_CLASSE                               			")
		query.Add(" WHERE SIPITEM IN (                                                 			")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

  		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_CID                               			")
		query.Add(" WHERE SIPANEXO IN (                                                			")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

  		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_ESPECIALIDA                               	")
		query.Add(" WHERE SIPANEXO IN (                                                			")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

  		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_IDADE                               			")
		query.Add(" WHERE SIPITEM IN (                                                 			")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM_GRAU                               			")
		query.Add(" WHERE SIPITEM IN (                                                 			")
		query.Add("			  SELECT SAI.HANDLE                                              	")
		query.Add("  			    FROM ANS_SIP_ANEXO ASA                                     	")
		query.Add("  			    JOIN ANS_SIP_ANEXO_ITEM SAI ON (ASA.HANDLE = SAI.SIPANEXO) 	")
		query.Add(" 			   WHERE ASA.HANDLE = :HANDLE)")
 		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
  		query.ExecSQL

		query.Active = False
		query.Clear
		query.Add("DELETE FROM ANS_SIP_ANEXO_ITEM WHERE SIPANEXO = :HANDLE")
		query.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsString
		query.ExecSQL

		Set query = Nothing

		bsShowMessage("Itens excluidos.", "I")
	End If
End Sub

Public Sub BOTAOVERIFICARDIVERGENCIAS_OnClick()
  Dim Obj As Object
  Dim vsMensagemRetorno As String
  Dim viRetorno As Long

  If VisibleMode Then
  	Set Obj = CreateBennerObject("BSINTERFACE0055.VerificarParametrizacao")
  	Obj.VerificarParametrizacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemRetorno)
  	Set Interface = Nothing
  Else
  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    "BSANS001", _
                                    "VerificarParametrizacao", _
                                    "SIP - Sistema de informações de Produtos - Verificar Divergências", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "ANS_SIP_ANEXO", _
                                    "SITUACAO", _
                                    "", _
                                    "", _
                                    "P", _
                                    False, _
                                    vsMensagemRetorno, _
                                    Null)

     Set Obj = Nothing
 	 If viRetorno = 0 Then
  		bsShowMessage("Processo enviado para execução no servidor!", "I")
 	 Else
  		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemRetorno, "I")
  	 End If
  End If

End Sub

Public Sub TABLE_AfterScroll()
	If CurrentQuery.FieldByName("ANEXO").AsString = "1" Then
    	RELATORIOANEXOI.Visible = True
  	Else
    	RELATORIOANEXOI.Visible = False
  	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("CONTABILIZAPARTOS").AsString = "S" Then
      Dim query As BPesquisa
	  Set query = NewQuery

      query.Active = False
	  query.Clear
	  query.Add("SELECT HANDLE FROM ANS_SIP_ANEXO    ")
	  query.Add(" WHERE CONTABILIZAPARTOS ='S'  ")
	  query.Active = True

	  If query.FieldByName("HANDLE").AsString <> "" Then
       bsShowMessage("Já existe registro com este campo marcado!", "I")
       CanContinue=False
	   Exit Sub
	  End If
	  Set query=Nothing
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCRIARITENS"
			BOTAOCRIARITENS_OnClick
		Case "BOTAODUPLICARANEXO"
			BOTAODUPLICARANEXO_OnClick
		Case "BOTAOEXCLUIRITENS"
			BOTAOEXCLUIRITENS_OnClick
		Case "BOTAOVERIFICARDIVERGENCIAS"
			BOTAOVERIFICARDIVERGENCIAS_OnClick
	End Select
End Sub
