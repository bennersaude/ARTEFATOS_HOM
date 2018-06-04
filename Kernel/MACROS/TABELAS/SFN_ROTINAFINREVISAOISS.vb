'HASH: 5D60D06E05EC2E93BDFF7943F63003F8
'Macro: SFN_ROTINAFINREVISAOISS
'#Uses "*bsShowMessage"
'#Uses "*ProcuraPrestador"

Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
  Dim ISSDLL As Object
  Dim vsDescricao As String

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT DESCRICAO")
  SQL.Add("FROM SFN_ROTINAFIN")
  SQL.Add("WHERE HANDLE = :HROTINAFIN")
  SQL.Active = False
  SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  SQL.Active = True

  vsDescricao = SQL.FieldByName("DESCRICAO").AsString

  Set SQL = Nothing

  If VisibleMode Then
	Set ISSDLL = CreateBennerObject("BSINTERFACE0029.Rotinas")
	ISSDLL.RevisaoISS_Cancelar(CurrentSystem, _
	   						  CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
	Dim vsMensagemErro As String
	Dim viRetorno As Integer

	Set ISSDLL = CreateBennerObject("BSSERVEREXEC.ProcessosServidor")
	viRetorno = ISSDLL.ExecucaoImediata(CurrentSystem, _
						   			    "SFNRECOLHIMENTO", _
										"Rotinas_RevisaoISS_Cancelar", _
										"Rotina de Revisão de ISS " + vsDescricao, _
										CurrentQuery.FieldByName("HANDLE").AsInteger, _
										"SFN_ROTINAFINREVISAOISS", _
										"SITUACAO", _
										"", _
										"", _
										"C", _
										False, _
										vsMensagemErro, _
										Null)
	If viRetorno = 0 Then
		bsShowMessage("Processo enviado para execução no servidor", "I")
	Else
		bsShowMessage("Erro ao enviar o processo para execução no servidor" + Chr(13) + vsMensagemErro, "I")
	End If
  End If

  Set ISSDLL = Nothing

  CurrentQuery.Active = False
  CurrentQuery.Active = True


  WriteAudit("C", HandleOfTable("SFN_ROTINAFINREVISAOISS"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Revisão de ISS - Cancelamento")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
 Dim ISSDLL As Object
  Dim SQL As Object
  Dim vsDescricao As String

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT DESCRICAO")
  SQL.Add("  FROM SFN_ROTINAFIN")
  SQL.Add(" WHERE HANDLE = :HANDLE")
  SQL.Active = False
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  SQL.Active = True

  vsDescricao = SQL.FieldByName("DESCRICAO").AsString

  Set SQL = Nothing

  If VisibleMode Then
	Set ISSDLL = CreateBennerObject("BSINTERFACE0029.Rotinas")
	ISSDLL.RevisaoISS_Processar(CurrentSystem, _
	   						    CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
	Dim vsMensagemErro As String
	Dim viRetorno As Integer

	Set ISSDLL = CreateBennerObject("BSSERVEREXEC.ProcessosServidor")
	viRetorno = ISSDLL.ExecucaoImediata(CurrentSystem, _
						   			    "SFNRECOLHIMENTO", _
										"Rotinas_RevisaoISS_Processar", _
										"Rotina de Revisão de ISS " + vsDescricao, _
										CurrentQuery.FieldByName("HANDLE").AsInteger, _
										"SFN_ROTINAFINREVISAOISS", _
										"SITUACAO", _
										"", _
										"", _
										"P", _
										False, _
										vsMensagemErro, _
										Null)
	If viRetorno = 0 Then
		bsShowMessage("Processo enviado para execução no servidor", "I")
	Else
		bsShowMessage("Erro ao enviar o processo para execução no servidor" + Chr(13) + vsMensagemErro, "I")
	End If
  End If

  Set ISSDLL = Nothing

  CurrentQuery.Active = False
  CurrentQuery.Active = True

  WriteAudit("P", HandleOfTable("SFN_ROTINAFINREVISAOISS"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Revisão de ISS - Processamento")
End Sub

'SMS 37423 - Filtrar os prestadores por Estado e Municipio, caso preenchidos

Public Sub ESTADO_OnChange()
  CurrentQuery.FieldByName("PRESTADOR").Clear

End Sub

Public Sub MUNICIPIO_OnChange()
  CurrentQuery.FieldByName("PRESTADOR").Clear

End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vsCPFNome As String
  Dim vCriterio As String

  ShowPopup = False

  vCriterio = ""
  If Not CurrentQuery.FieldByName("ESTADO").IsNull Then
    vCriterio = " AND ESTADOPAGAMENTO = " + CurrentQuery.FieldByName("ESTADO").AsString
    If Not CurrentQuery.FieldByName("MUNICIPIO").IsNull Then
      vCriterio = vCriterio + " AND MUNICIPIOPAGAMENTO = " + CurrentQuery.FieldByName("MUNICIPIO").AsString
    End If
  End If

  If (IsNumeric(PRESTADOR.LocateText)) Then
    vsCPFNome = "C"
  Else
    vsCPFNome = "N"
  End If

  vHandle = ProcuraPrestador(vsCPFNome, "T", PRESTADOR.LocateText, vCriterio)

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
End Sub

Public Sub TABLE_AfterScroll()

  BOTAOPROCESSAR.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString = "1"
  BOTAOCANCELAR.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString = "5"


End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
		CanContinue = False
		bsShowMessage("Rotina não está aberta", "E")
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
		CanContinue = False
		bsShowMessage("Rotina não está aberta", "E")
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOPROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	ElseIf CommandID = "BOTAOCANCELAR" Then
		BOTAOCANCELAR_OnClick
	End If
End Sub
