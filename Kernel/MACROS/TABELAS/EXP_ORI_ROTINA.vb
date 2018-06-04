'HASH: 0E39774AD08A8E65DD51E29E7B099548
Option Explicit
'#Uses "*bsShowMessage"

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os dados não podem estar em edição!", "E")
    Exit Sub
  End If

  Dim interface As Object
  Dim vsCodigoEMS As String
  Dim qCodigoEMS As Object

  Set qCodigoEMS = NewQuery

  qCodigoEMS.Add("SELECT CODIGOEMS FROM EMPRESAS WHERE HANDLE = :HANDLE")
  qCodigoEMS.ParamByName("HANDLE").AsInteger = CurrentCompany
  qCodigoEMS.Active = True

  Set interface = CreateBennerObject("Benner.Saude.Orizon.Exportacao.ExportacaoOrizon")
  interface.ExportarDadosOrizon(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("CODIGO").AsInteger, qCodigoEMS.FieldByName("CODIGOEMS").AsString, RetornaTabelasSelecionadas, CurrentQuery.FieldByName("DIRETORIOEXPORTACAO").AsString, CurrentQuery.FieldByName("DIRETORIOACESSO").AsString)

  Set qCodigoEMS = Nothing

  RefreshNodesWithTable("EXP_ORI_ROTINA")

End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("ORIGEM").AsString = "M"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If RetornaTabelasSelecionadas = "" Then
    bsShowMessage("Ao menos uma tabela deve ser selecionada para exportação!", "E")
    CanContinue = False
  End If
End Sub

Public Sub TABLE_NewRecord()
  Dim viChave As Long

  Dim qParamTISS As Object
  Set qParamTISS = NewQuery

  CurrentQuery.FieldByName("STATUS").AsInteger = 1

  NewCounter2("EXP_ORI_ROTINA", 0, 1, viChave)
  CurrentQuery.FieldByName("CODIGO").AsInteger = viChave

  qParamTISS.Add("SELECT DIRETORIOEXPORTACAOORIZON, DIRETORIOACESSOORIZON FROM TIS_PARAMETROS")
  qParamTISS.Active = True

  CurrentQuery.FieldByName("DIRETORIOEXPORTACAO").AsString = qParamTISS.FieldByName("DIRETORIOEXPORTACAOORIZON").AsString
  CurrentQuery.FieldByName("DIRETORIOACESSO").AsString = qParamTISS.FieldByName("DIRETORIOACESSOORIZON").AsString

  Set qParamTISS = Nothing

End Sub

Function RetornaTabelasSelecionadas As String
  RetornaTabelasSelecionadas = ""

  If CurrentQuery.FieldByName("TABELAEMS").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "EMS|"
  End If

  If CurrentQuery.FieldByName("TABELACRD").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "CRD|"
  End If

  If CurrentQuery.FieldByName("TABELAPRS").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PRS|"
  End If

  If CurrentQuery.FieldByName("TABELAPCO").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PCO|"
  End If

  If CurrentQuery.FieldByName("TABELAPCA").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PCA|"
  End If

  If CurrentQuery.FieldByName("TABELAPES").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PES|"
  End If

  If CurrentQuery.FieldByName("TABELAPDM").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PDM|"
  End If

  If CurrentQuery.FieldByName("TABELAESP").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "ESP|"
  End If

  If CurrentQuery.FieldByName("TABELAPDE").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PDE|"
  End If

  If CurrentQuery.FieldByName("TABELAUSU").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "USU|"
  End If

  If CurrentQuery.FieldByName("TABELAEMP").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "EMP|"
  End If

  If CurrentQuery.FieldByName("TABELAPLN").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PLN|"
  End If

  If CurrentQuery.FieldByName("TABELADUS").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "DUS|"
  End If

  If CurrentQuery.FieldByName("TABELADPL").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "DPL|"
  End If

  If CurrentQuery.FieldByName("TABELADEM").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "DEM|"
  End If

  If CurrentQuery.FieldByName("TABELADUG").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "DUG|"
  End If
  If CurrentQuery.FieldByName("TABELADPG").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "DPG|"
  End If

  If CurrentQuery.FieldByName("TABELADEG").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "DEG|"
  End If

  If CurrentQuery.FieldByName("TABELAGRP").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "GRP|"
  End If

  If CurrentQuery.FieldByName("TABELAPGR").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PGR|"
  End If

  If CurrentQuery.FieldByName("TABELAINP").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "INP|"
  End If

  If CurrentQuery.FieldByName("TABELAPCP").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PCP|"
  End If

  If CurrentQuery.FieldByName("TABELAPLI").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PLI|"
  End If

  If CurrentQuery.FieldByName("TABELAPIN").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PIN|"
  End If

  If CurrentQuery.FieldByName("TABELAPDMA").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PDMA|"
  End If

  If CurrentQuery.FieldByName("TABELAERP").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "ERP|"
  End If

  If CurrentQuery.FieldByName("TABELARUS").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "RUS|"
  End If

  If CurrentQuery.FieldByName("TABELAGPC").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "GPC|"
  End If

  If CurrentQuery.FieldByName("TABELAPTT").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "PTT|"
  End If

  If CurrentQuery.FieldByName("TABELAOPDM").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "OPDM|"
  End If
  If CurrentQuery.FieldByName("TABELAOPRG").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "OPRG|"
  End If
  If CurrentQuery.FieldByName("TABELAOPDT").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "OPDT|"
  End If
  If CurrentQuery.FieldByName("TABELAOPDF").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "OPDF|"
  End If
  If CurrentQuery.FieldByName("TABELAOPU").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "OPU|"
  End If
  If CurrentQuery.FieldByName("TABELAOPEE").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "OPEE|"
  End If
  If CurrentQuery.FieldByName("TABELAOPLI").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "OPLI|"
  End If
  If CurrentQuery.FieldByName("TABELAVPR").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "VPR|"
  End If
  If CurrentQuery.FieldByName("TABELAVUF").AsString = "S" Then
    RetornaTabelasSelecionadas = RetornaTabelasSelecionadas + "VUF|"
  End If

End Function
