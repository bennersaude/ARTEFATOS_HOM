'HASH: 57905E53913D68EFDB89CDCE8475F243
'Macro: SAM_ROTAVISOREAJUSTE


Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.FieldByName("processado").Value = "S" Then

    Dim MODULO As Object
    Dim BENEFICIARIO As Object
    Dim PROCESSA As Object
    Dim FAMILIA As Object
    Set MODULO = NewQuery
    Set BENEFICIARIO = NewQuery
    Set PROCESSA = NewQuery
    Set FAMILIA = NewQuery

    If MsgBox("Todos os dados já cadastrados serão apagados!" + Chr(13) + "Deseja Continuar?", 4) = vbYes Then

      If Not InTransaction Then StartTransaction

      MODULO.Active = False
      MODULO.Clear
      MODULO.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_MOD")
      MODULO.Add("WHERE HANDLE IN (SELECT ROTMOD.HANDLE")
      MODULO.Add("  		FROM SAM_ROTAVISOREAJUSTE ROT,")
      MODULO.Add("       		SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
      MODULO.Add("       		SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM,")
      MODULO.Add("       		SAM_ROTAVISOREAJUSTE_BENEF ROTBENEF,")
      MODULO.Add("       		SAM_ROTAVISOREAJUSTE_MOD ROTMOD")
      MODULO.Add(" 			WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
      MODULO.Add("   			And ROTCONT.HANDLE = ROTFAM.ROTINAAVISOCONTRATO")
      MODULO.Add("   			And ROTFAM.HANDLE = ROTBENEF.ROTINAAVISOFAMILIA")
      MODULO.Add("   			And ROTBENEF.HANDLE = ROTMOD.ROTINAAVISOBENEF")
      MODULO.Add("   			And ROT.HANDLE = :HANDLEROTINA)")
      MODULO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      MODULO.ExecSQL

      BENEFICIARIO.Active = False
      BENEFICIARIO.Clear
      BENEFICIARIO.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_BENEF")
      BENEFICIARIO.Add("WHERE HANDLE IN (SELECT ROTBENEF.HANDLE")
      BENEFICIARIO.Add("  			FROM SAM_ROTAVISOREAJUSTE ROT,")
      BENEFICIARIO.Add("       			SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
      BENEFICIARIO.Add("       			SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM,")
      BENEFICIARIO.Add("       			SAM_ROTAVISOREAJUSTE_BENEF ROTBENEF")
      BENEFICIARIO.Add(" 			WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
      BENEFICIARIO.Add("   				And ROTCONT.HANDLE = ROTFAM.ROTINAAVISOCONTRATO")
      BENEFICIARIO.Add("   				And ROTFAM.HANDLE = ROTBENEF.ROTINAAVISOFAMILIA")
      BENEFICIARIO.Add("   				And ROT.HANDLE = :HANDLEROTINA)")
      BENEFICIARIO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      BENEFICIARIO.ExecSQL

      FAMILIA.Active = False
      FAMILIA.Clear
      FAMILIA.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_FAMILIA")
      FAMILIA.Add("WHERE HANDLE IN (SELECT ROTFAM.HANDLE ")
      FAMILIA.Add("  			FROM SAM_ROTAVISOREAJUSTE ROT,")
      FAMILIA.Add("       			SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
      FAMILIA.Add("       			SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM")
      FAMILIA.Add(" 			WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
      FAMILIA.Add("   				And ROTCONT.HANDLE = ROTFAM.ROTINAAVISOCONTRATO")
      FAMILIA.Add("   				And ROT.HANDLE = :HANDLEROTINA)")
      FAMILIA.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      FAMILIA.ExecSQL


      PROCESSA.Active = False
      PROCESSA.Clear
      PROCESSA.Add("UPDATE SAM_ROTAVISOREAJUSTE")
      PROCESSA.Add("   Set PROCESSADO = 'N',")
      PROCESSA.Add("       OCORRENCIAS = ''")
      PROCESSA.Add(" WHERE HANDLE = :HANDLEROTINA")
      PROCESSA.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      PROCESSA.ExecSQL

      If InTransaction Then Commit

      WriteAudit("C", HandleOfTable("SAM_ROTAVISOREAJUSTE"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Aviso Reajuste - Cancelamento")

      RefreshNodesWithTable("SAM_ROTAVISOREAJUSTE")

    Else

      CurrentQuery.Cancel
      CONCONTINUE = False

    End If


    Set MODULO = Nothing
    Set BENEFICARIO = Nothing
    Set PROCESSA = Nothing
    Set FAMILIA = Nothing
  Else
    MsgBox("Rotina ainda não processada!!!")
  End If
End Sub


Public Sub BOTAOGERAR_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SamAvisoReajuste.Geral")
  interface.GerarContrato(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing
End Sub

Public Sub BOTAOIMPRIMRINIVFAMIL_OnClick()
  Dim RelatorioHandle As Long
  Dim HandleRotCanc As Long
  Dim QueryBuscaHandleRelatorio As Object

  Set QueryBuscaHandleRelatorio = NewQuery

  QueryBuscaHandleRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'BEN019ANS'")
  QueryBuscaHandleRelatorio.Active = False
  QueryBuscaHandleRelatorio.Active = True

  RelatorioHandle = QueryBuscaHandleRelatorio.FieldByName("HANDLE").AsInteger

  Set QueryBuscaHandleRelatorio = Nothing

  HandleRotCanc = CurrentQuery.FieldByName("HANDLE").AsInteger

  If MsgBox("Emitir aviso de reajuste Individual/Familiar?", vbYesNo) = vbYes Then
    ReportPreview(RelatorioHandle, "A.HANDLE=" + Str(HandleRotCanc), False, False)

  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SamAvisoReajuste.Geral")
  interface.Aviso(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing

  WriteAudit("P", HandleOfTable("SAM_ROTAVISOREAJUSTE"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina de Aviso Reajuste - Processamento")

  RefreshNodesWithTable("SAM_ROTAVISOREAJUSTE")

End Sub

Public Sub TABLE_AfterInsert()

  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser
  CurrentQuery.FieldByName("COMPETENCIAFINAL").Value = DateAdd("m", 1, CurrentQuery.FieldByName("COMPETENCIAATUAL").Value)

End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("processado").AsString = "S" Then
    BOTAOPROCESSAR.Caption = "Reprocessar"
  Else
    BOTAOPROCESSAR.Caption = "Processar"
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim MODULO As Object
  Dim BENEFICIARIO As Object
  Dim CONTRATO As Object
  Dim FAMILIA As Object
  Set MODULO = NewQuery
  Set BENEFICIARIO = NewQuery
  Set CONTRATO = NewQuery
  Set FAMILIA = NewQuery

  MODULO.Active = False
  MODULO.Clear
  MODULO.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_MOD")
  MODULO.Add("WHERE HANDLE IN (SELECT ROTMOD.HANDLE")
  MODULO.Add("  		FROM SAM_ROTAVISOREAJUSTE ROT,")
  MODULO.Add("       		SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
  MODULO.Add("       		SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM,")
  MODULO.Add("       		SAM_ROTAVISOREAJUSTE_BENEF ROTBENEF,")
  MODULO.Add("       		SAM_ROTAVISOREAJUSTE_MOD ROTMOD")
  MODULO.Add(" 			WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
  MODULO.Add("   			And ROTCONT.HANDLE = ROTFAM.ROTINAAVISOCONTRATO")
  MODULO.Add("   			And ROTFAM.HANDLE = ROTBENEF.ROTINAAVISOFAMILIA")
  MODULO.Add("   			And ROTBENEF.HANDLE = ROTMOD.ROTINAAVISOBENEF")
  MODULO.Add("   			And ROT.HANDLE = :HANDLEROTINA)")
  MODULO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  MODULO.ExecSQL

  BENEFICIARIO.Active = False
  BENEFICIARIO.Clear
  BENEFICIARIO.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_BENEF")
  BENEFICIARIO.Add("WHERE HANDLE IN (SELECT ROTBENEF.HANDLE")
  BENEFICIARIO.Add("  			FROM SAM_ROTAVISOREAJUSTE ROT,")
  BENEFICIARIO.Add("       			SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
  BENEFICIARIO.Add("       			SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM,")
  BENEFICIARIO.Add("       			SAM_ROTAVISOREAJUSTE_BENEF ROTBENEF")
  BENEFICIARIO.Add(" 			WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
  BENEFICIARIO.Add("   				And ROTCONT.HANDLE = ROTFAM.ROTINAAVISOCONTRATO")
  BENEFICIARIO.Add("   				And ROTFAM.HANDLE = ROTBENEF.ROTINAAVISOFAMILIA")
  BENEFICIARIO.Add("   				And ROT.HANDLE = :HANDLEROTINA)")
  BENEFICIARIO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  BENEFICIARIO.ExecSQL

  FAMILIA.Active = False
  FAMILIA.Clear
  FAMILIA.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_FAMILIA")
  FAMILIA.Add(" WHERE HANDLE IN (SELECT ROTFAM.HANDLE ")
  FAMILIA.Add("                    FROM SAM_ROTAVISOREAJUSTE ROT,")
  FAMILIA.Add("                         SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT,")
  FAMILIA.Add("                         SAM_ROTAVISOREAJUSTE_FAMILIA ROTFAM")
  FAMILIA.Add("                   WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
  FAMILIA.Add("                     And ROTCONT.HANDLE = ROTFAM.ROTINAAVISOCONTRATO")
  FAMILIA.Add("   	                And ROT.HANDLE = :HANDLEROTINA)")
  FAMILIA.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  FAMILIA.ExecSQL

  CONTRATO.Active = False
  CONTRATO.Clear
  CONTRATO.Add("DELETE FROM SAM_ROTAVISOREAJUSTE_CONTRATO")
  CONTRATO.Add(" WHERE HANDLE IN (SELECT ROTCONT.HANDLE ")
  CONTRATO.Add("                    FROM SAM_ROTAVISOREAJUSTE ROT,")
  CONTRATO.Add("                         SAM_ROTAVISOREAJUSTE_CONTRATO ROTCONT")
  CONTRATO.Add("                   WHERE ROT.HANDLE = ROTCONT.ROTINAVISO")
  CONTRATO.Add("                     And ROT.HANDLE = :HANDLEROTINA)")
  CONTRATO.ParamByName("HANDLEROTINA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  CONTRATO.ExecSQL

  Set MODULO = Nothing
  Set BENEFICARIO = Nothing
  Set CONTTRATO = Nothing
  Set FAMILIA = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  CurrentQuery.FieldByName("COMPETENCIAATUAL").Value = DateAdd("m", -1, CurrentQuery.FieldByName("COMPETENCIAFINAL").Value)

End Sub
