'HASH: C8C647264C92DCBF9216E9FA364B5D53
Option Explicit

Public Sub ATUALIZAREGRAPARC_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("PARCELAMENTO.Rotinas")
  interface.AtualizaRegraParcelamento(CurrentSystem,)
  Set interface =Nothing
End Sub

Public Sub CANCELARDMED_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0081", _
					   "Cancelar Dmed",  _
					   0, _
					   100, _
					   180, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  Set INTERFACE0002 = Nothing
End Sub

Public Sub CONSULTACONTACORR_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("SFNContaCorrente.Consulta")
    interface.Executar(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub CONSULTAFINANCEIRA_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0110", _
					   "Consulta financeira",  _
					   0, _
					   480, _
					   350, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  Set INTERFACE0002 = Nothing
End Sub

Public Sub CONSULTAGERENCIAL_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("SfnGerencial.Rotinas")
  interface.Exec(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub DEMONSTRATIVOINDIVID_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("rotarq.rotinas")
  interface.ArqDemonstrativoIndividual(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub DEMONSTRATIVOPAG_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("RotArq.Rotinas")
  interface.ArqDemonstrativoPagto(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub DEMONSTRATIVOREC_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("RotArq.Rotinas")
  interface.ArqDemonstrativoRecto(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub EXPORTAARQUIVOSIC_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("ROTARQ.Rotinas")
  interface.ArqContab(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub IMPORTAFATURA_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("ROTARQ.Rotinas")
  interface.ImportacaoFaturas(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub xINSS_OnClick()'pode excluir
  Dim interface As Object
  Set interface =CreateBennerObject("SamINSS.INSS")
  interface.Faturar(CurrentSystem)
  Set interface =Nothing
End Sub


Public Sub EXCLUSAODOCUMENTO_OnClick()
  Dim SQL As Object
  Set SQL =NewQuery
  SQL.Clear
  SQL.Add("SELECT G.HANDLE FROM Z_GRUPOS G WHERE G.HANDLE IN (SELECT U.GRUPO FROM Z_GRUPOUSUARIOS U WHERE U.HANDLE = :USUARIO)")
  SQL.ParamByName("USUARIO").Value =CurrentUser
  SQL.Active =True
  If Not SQL.EOF Then
    Dim interface As Object
    Set interface =CreateBennerObject("Financeiro.ExclusaoDoc")
    interface.Exec(CurrentSystem)
  End If
  Set interface =Nothing
End Sub

Public Sub GERARDADOSIR_OnClick()
  Dim mensagemErro As String
  Dim retorno As Integer

  Dim interface As Object
  Set interface =CreateBennerObject("BSDMED.GERACAODADOSIR")

  Dim qTabela As BPesquisa
  Set qTabela = NewQuery

  qTabela.Active = False
  qTabela.Add("SELECT HANDLE FROM Z_TABELAS WHERE NOME = 'SFN_DMEDANOCALENDARIO'")
  qTabela.Active = True

  If (CurrentTable = qTabela.FieldByName("HANDLE").AsInteger) Then
  	If CurrentQuery.Active Then
    	retorno = interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, mensagemErro)
    Else
    	retorno = interface.Processar(CurrentSystem, 0, mensagemErro)
    End If
  Else
	retorno = interface.Processar(CurrentSystem, 0, mensagemErro)
  End If

 If retorno = 1 Then
  	MsgBox(mensagemErro)
  End If

  qTabela.Active = False
  Set interface =Nothing
  Set qTabela = Nothing
End Sub

Public Sub MODULE_BeforeNodeShow(ByVal NodeFullPath As String, CanShow As Boolean)
	If (NodeFullPath = "DOTACAOORCAMENTARIA" Or NodeFullPath = "7.9 Tabelas do Financeiro|7.20DOTACAOORCAMENTARIA") And CurrentModuleName = "Controle Financeiro" Then
	  Dim SQL As Object
	  Set SQL =NewQuery
	  SQL.Clear
	  SQL.Add("SELECT 1 FROM SFN_PARAMETROSFIN WHERE CONTROLADOTORC = 1")
	  SQL.Active =True
	  If Not SQL.EOF Then
		CanShow = False
	  End If
	End If
End Sub

Public Sub RECLASSIFICAGUIAPAG_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("SamPEG.PROCESSAR")
  interface.Reclassificacao(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub RECONTABILIZACAO_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("Financeiro.Geral")
  interface.Recontabilizacao(CurrentSystem)
  Set interface =Nothing

End Sub

Public Sub VERIFICABENEFICIARIO_OnClick()

  If MsgBox("Confirma verificação de todos beneficiários que devem ter conta financeira?",vbYesNo +vbDefaultButton2)=vbYes Then

  Dim interface As Object
  Dim SQL As Object
  Dim QUANT As Integer
  Dim c As Integer


  Set SQL =NewQuery
  Set interface =CreateBennerObject("FINANCEIRO.ContaFin")
  SQL.Add("SELECT COUNT(HANDLE) QUANT FROM SAM_BENEFICIARIO WHERE EHTITULAR='S'")
  SQL.Active =True
  QUANT =SQL.FieldByName("QUANT").AsInteger
  Progress.Init("Beneficiários 1/" +Str(QUANT),0,100)
  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_BENEFICIARIO WHERE EHTITULAR='S'")
  SQL.Active =True
  c =0
  While Not SQL.EOF
    If interface.Cadastro(CurrentSystem,SQL.FieldByName("HANDLE").AsInteger,1,0)<0 Then
      MsgBox "Erro ao criar conta financeira"
    End If
    Progress.Value =Progress.Value +1
    c =c +1
    If Progress.Value =100 Then
      Progress.Init("Beneficiários " +Str(c)+"/" +Str(QUANT),0,100)
    End If
    SQL.Next
  Wend

  Progress.Visible =False
  Set SQL =Nothing
  Set interface =Nothing
 End If
End Sub

Public Sub VERIFICAPESSOA_OnClick()
  If MsgBox("Confirma verificação de todas pessoas que devem ter conta financeira?",vbYesNo +vbDefaultButton2)=vbYes Then

  Dim Interface As Object
  Dim SQL As Object
  Dim QUANT As Integer
  Dim c As Integer

  Set SQL =NewQuery
  Set Interface =CreateBennerObject("FINANCEIRO.ContaFin")
  SQL.Add("SELECT COUNT(HANDLE) QUANT FROM SFN_PESSOA")
  SQL.Active =True

  QUANT =SQL.FieldByName("QUANT").AsInteger
  Progress.Init("Pessoa 1/" +Str(QUANT),0,100)

  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SFN_PESSOA")
  SQL.Active =True
  While Not SQL.EOF
    If Interface.Cadastro(CurrentSystem,SQL.FieldByName("HANDLE").AsInteger,3,0)<0 Then
      MsgBox "Erro ao criar conta financeira"
    End If
    Progress.Value =Progress.Value +1
    c =c +1
    If Progress.Value =100 Then
      Progress.Init("Pessoa " +Str(c)+"/" +Str(QUANT),0,100)
    End If
    SQL.Next
  Wend

  Progress.Visible =False
  Set SQL =Nothing
  Set Interface =Nothing
  End If

End Sub

Public Sub VERIFICAPRESTADOR_OnClick()
  If MsgBox("Confirma verificação de todos prestadores que devem ter conta financeira?",vbYesNo +vbDefaultButton2)=vbYes Then

  Dim Interface As Object
  Dim SQL As Object
  Dim QUANT As Integer
  Dim c As Integer

  Set SQL =NewQuery
  Set Interface =CreateBennerObject("FINANCEIRO.ContaFin")
  SQL.Add("SELECT COUNT(HANDLE) QUANT FROM SAM_PRESTADOR")
  SQL.Active =True

  QUANT =SQL.FieldByName("QUANT").AsInteger
  Progress.Init("Prestador 1/" +Str(QUANT),0,100)

  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_PRESTADOR")
  SQL.Active =True
  While Not SQL.EOF
    If Interface.Cadastro(CurrentSystem,SQL.FieldByName("HANDLE").AsInteger,2,0)<0 Then
      MsgBox "Erro ao criar conta financeira"
    End If
    Progress.Value =Progress.Value +1
    c =c +1
    If Progress.Value =100 Then
      Progress.Init("Prestador " +Str(c)+"/" +Str(QUANT),0,100)
    End If
    SQL.Next
  Wend

  Progress.Visible =False
  Set SQL =Nothing
  Set Interface =Nothing
  End If
End Sub
Public Sub PROCESSARDMED_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0070", _
					   "Processar Dmed",  _
					   0, _
					   270, _
					   280, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  Set INTERFACE0002 = Nothing

End Sub

Public Sub RETIFICARDMED_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0071", _
					   "Retificar Dmed",  _
					   0, _
					   100, _
					   180, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  Set INTERFACE0002 = Nothing
End Sub

Public Sub GERARARQUIVODMED_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0072", _
					   "Gerar arquivo Dmed",  _
					   0, _
					   100, _
					   180, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  Set INTERFACE0002 = Nothing
End Sub

Public Sub MONITORESOCIAL_OnClick()
  Dim Monitor As Object
  Set Monitor = CreateManagedObject("Benner.Saude.eSocial.Business", "Benner.Saude.eSocial.Business.Formularios.Interface")
  Monitor.ExibirMonitorEventosESocial()
  Set Monitor = Nothing
End Sub
