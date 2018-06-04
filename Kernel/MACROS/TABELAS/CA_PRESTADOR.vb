'HASH: 426D2153DB228E6C2E2298E25068A58B
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOALTERACOES_OnClick()
  Dim INTERFACE0002 As Object
  Dim vSMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  INTERFACE0002.Exec(CurrentSystem, _
                       1, _
                       "TV_FORM_CA_PRESTADOR_ALTERACAO", _
                       "Alterações", _
                       0, _
                       400, _
                       420, _
                       False, _
                       vSMensagem, _
                       vcContainer)

  Set INTERFACE0002 = Nothing
End Sub

Public Sub BOTAOALTERACOESENDERECO_OnClick()
  Dim dll As Object
  Set dll = CreateBennerObject("CA032.ALTERAPRESTADOR")
  dll.ExibirAlteracoesEndereco(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set dll = Nothing
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim dll As Object
  If(CurrentQuery.FieldByName("SITUACAO").AsString = "C")Then
  	bsShowMessage("Esta solicitação está cancelada", "I")
  Else
    If(CurrentQuery.FieldByName("SITUACAO").AsString = "P")Then
  	  bsShowMessage("Esta solicitação já foi processada. Não é possível cancelar!", "I")
    Else
	  Set dll = CreateBennerObject("CA032.ALTERAPRESTADOR")
      dll.CancelarAlteracoes(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      bsShowMessage("Solicitação cancelada com sucesso!", "I")
      FinalizarAlteracaoCadastral CurrentQuery.FieldByName("HANDLE").AsInteger, "Cancelada"
      RefreshNodesWithTable("CA_PRESTADOR")
    End If
  End If
  Set dll = Nothing
End Sub

Public Sub BOTAODETALHES_OnClick()
  Dim dll As Object
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM SAM_PRESTADOR WHERE PRESTADOR = :HANDLE")
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").Value
  sql.Active = True
  Set dll = CreateBennerObject("CA005.ConsultaPrestador")
  dll.Info(CurrentSystem, sql.FieldByName("HANDLE").AsInteger, 0)
  Set dll = Nothing
End Sub

Public Sub BOTAOEFETIVAR_OnClick()
  Dim dll As Object
  Dim sql As Object
  If CurrentQuery.FieldByName("SITUACAO").AsString = "C" Then
	bsShowMessage("Solicitação está cancelada. Não é possível efetivar alterações", "I")
  Else
    If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
      bsShowMessage("Alterações já efetivadas. Operação cancelada", "I")
    Else
      Set dll = CreateBennerObject("CA032.ALTERAPRESTADOR")
      dll.EfetivarAlteracoes(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    End If
  End If
  If CurrentQuery.FieldByName("SITUACAO").AsString = "A" Then
	bsShowMessage("Solicitação efetivada com sucesso", "I")
	FinalizarAlteracaoCadastral CurrentQuery.FieldByName("HANDLE").AsInteger, "Efetivada"
  End If
  RefreshNodesWithTable("CA_PRESTADOR")
  Set dll = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  BOTAODETALHES.Visible = True
  BOTAOALTERACOES.Visible = True
  BOTAOEFETIVAR.Visible = True
  BOTAOCANCELAR.Visible = True
  BOTAOALTERACOESENDERECO.Visible = True

  If CurrentQuery.FieldByName("SITUACAO").AsString = "A" Then
	BOTAODETALHES.Visible = False
	BOTAOALTERACOES.Visible = True
	BOTAOEFETIVAR.Visible = True
	BOTAOCANCELAR.Visible = True
	BOTAOALTERACOESENDERECO.Visible = False
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "C" Then
    BOTAODETALHES.Visible = False
	BOTAOALTERACOES.Visible = False
	BOTAOEFETIVAR.Visible = False
	BOTAOCANCELAR.Visible = False
	BOTAOALTERACOESENDERECO.Visible = False
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
	BOTAODETALHES.Visible = False
	BOTAOALTERACOES.Visible = False
	BOTAOEFETIVAR.Visible = False
	BOTAOCANCELAR.Visible = False
	BOTAOALTERACOESENDERECO.Visible = False
  End If
  SessionVar("hAlteracoes") = CurrentQuery.FieldByName("HANDLE").AsString
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
	Case "BOTAOALTERACOES"
		BOTAOALTERACOES_OnClick
	Case "BOTAOALTERACOESENDERECO"
		BOTAOALTERACOESENDERECO_OnClick
	Case "BOTAOCANCELAR"
		BOTAOCANCELAR_OnClick
	Case "BOTAOCANCELAR"
		BOTAOCANCELAR_OnClick
	Case "BOTAODETALHES"
		BOTAODETALHES_OnClick
	Case "BOTAOEFETIVAR"
		BOTAOEFETIVAR_OnClick
  End Select
End Sub

Public Sub FinalizarAlteracaoCadastral(handleAlterCad As Long, acao As String)
	On Error GoTo Erro
		Dim componente As CSBusinessComponent

		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.AlteracaoCadastral.CaPrestadorBLL, Benner.Saude.Prestadores.Business")
	    componente.AddParameter(pdtInteger, handleAlterCad)
	  	componente.AddParameter(pdtString, acao)
	  	componente.Execute("FinalizarAlteracaoCadastral")

	    Set componente = Nothing
	    Exit Sub
	Erro:
		bsShowMessage(Err.Description, "I")
		Set componente = Nothing
		Exit Sub
End Sub

