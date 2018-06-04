'HASH: 8539758A60767178C401D054C5F04183

Option Explicit

Public Sub CONSULTAESPEPREST_OnClick()
	Dim OLESamObject As Object
	Set OLESamObject =CreateBennerObject("SamCorresp.Controle")
	OLESamObject.Rotina(CurrentSystem)
	Set OLESamObject =Nothing
End Sub

Public Sub ATENDIMENTO_OnClick()


   'If MsgBox("Entrar na Central de Atendimento??",vbYesNo) =vbYes  Then

      Dim Interface As Object


      Set Interface =CreateBennerObject("CA001.CATEND")
      On Error GoTo A
      'MsgBox CurrentSystem.UserNickName

      Interface.Exec(CurrentSystem)
      Set Interface =Nothing
      Exit Sub
      A : Interface.Exec(CurrentSystem)
      Set Interface =Nothing
    'Else
    ' Dim INTERFACE2 As Object
    ' Set INTERFACE2=CreateBennerObject("cadlls.dll")
    ' INTERFACE2.Exec
    ' Set INTERFACE2=Nothing
    'End If
End Sub

Public Sub AUTORIZACAOSIMPLES_OnClick()
  Dim Interface As Object
  Set Interface =CreateBennerObject("SamAutoDigit.Formulario")
  Interface.Simples(CurrentSystem)
  Set Interface =Nothing
End Sub

Public Sub CONSULTABENEFICIARIO_OnClick()
  'Set interface =CreateBennerObject("CA010.ConsultaBeneficiario")
  'Alterado SMS 108068 - Crislei.Sorrilha 18/02/2009 -
  'Separação da Interface da regra de negocio para consulta de Beneficiários
  Dim Interface As Object
  Set Interface =CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")
  Interface.filtro(CurrentSystem,1,"")
End Sub

Public Sub CONSULTAEVENTOPACOTE_OnClick()
	Dim OLESamObject As Object
	Set OLESamObject =CreateBennerObject("SamConsulta.Consulta")
	OLESamObject.EventosPacote(CurrentSystem,-1,-1,1950 -1 -1)
	Set OLESamObject =Nothing
End Sub

Public Sub CONSULTAEVENTOPREST_OnClick()
	Dim OLESamObject As Object
	Set OLESamObject =CreateBennerObject("BSINTERFACE0020.Consulta")
	OLESamObject.Abrir(CurrentSystem)
	Set OLESamObject =Nothing
End Sub

Public Sub CONSULTALOGATEND_OnClick()
  Dim Interface As Object
  Set Interface = CreateBennerObject("CA028.ConsultaLogAtend")
  Interface.Exec(CurrentSystem)
  Set Interface = Nothing
End Sub

Public Sub CONSULTAPRECOEVENTO_OnClick()
  'Dim Interface As Object
  'Set Interface =CreateBennerObject("PRECO.PEGAPRECO")
  'Interface.ComandoPrecoEvento(CurrentSystem)
  'Set Interface =Nothing

  Dim Interface As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String
  Dim vvContainer As CSDContainer

  Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
  viRetorno = Interface.Exec(CurrentSystem, _
                             1, _
                             "TV_FORM0004", _
                             "Consulta Preço do Evento", _
                             0, _
                             740, _
                             640, _
                             False, _
                             vsMensagem, _
                             vvContainer)

  Set Interface =Nothing

End Sub

Public Sub CONSULTAPRESTADOR_OnClick()
	Dim Interface As Object
	Dim vsMensagem As String
	Dim vlHPrestador As Long

	Set interface = CreateBennerObject("BSINTERFACE0001.BuscaPrestador")
	interface.Abrir(CurrentSystem, vsMensagem, 1, "", "T", vlHPrestador)

	Set interface =Nothing
End Sub

Public Sub CONSULTAREEMBOLSO_OnClick()
  Dim SQL As Object

  Set SQL =NewQuery

  SQL.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'ATE028'")
  SQL.Active =True

  ReportPreview(SQL.FieldByName("HANDLE").Value,"",True,False)

  Set SQL =Nothing
End Sub

Public Sub CONSULTAREVENTO_OnClick()
  Dim interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Set Interface =CreateBennerObject("Procura.Procurar")

  vColunas ="ESTRUTURA|Z_DESCRICAO|NIVELAUTORIZACAO|JUSTIFICATIVA|EXAMEPREOPERATORIO|EXAMEPOSOPERATORIO"

  vCriterio ="*SAM_TGE.ULTIMONIVEL = 'S'"

  vCampos ="Evento|Descrição|Nível|Justific|Exame Pré|Exame Pós"

  Interface.Exec(CurrentSystem,"SAM_TGE",vColunas,2,vCampos,vCriterio,"Tabela Geral de Eventos",False,"","CA011.CONSULTATGE")

  Set Interface =Nothing
End Sub

Public Sub DIGITABENEFICIARIO_OnClick()
  Dim dllBSInterface001 As Object

  Set dllBSInterface001 = CreateBennerObject("BSINTERFACE0011.DigitarBeneficiario")

  dllBSInterface001.Exec(CurrentSystem, _
						 0, _
						 0, _
						 0)

  Set dllBSInterface001 = Nothing
End Sub

Public Sub DIGITACAOAUTORIZACAO_OnClick()
      Dim Interface As Object
      Set Interface =CreateBennerObject("ca014.Autorizacao")
      Interface.Exec(CurrentSystem,0,0,0)
      Set Interface =Nothing
End Sub

Public Sub FAXAUTORIZACAO_OnClick()
Dim Interface As Object
Set Interface =CreateBennerObject("CA030.Rotinas")
Interface.Executar(CurrentSystem,)
Set Interface =Nothing
End Sub

Public Sub DIGITACAONOVAAUTORIZ_OnClick()
  Dim Interface As Object
  UserVar("BENEFICIARIO") =""
  Set Interface =CreateBennerObject("CA043.Autorizacao")
  Interface.Executar(CurrentSystem,0,0,0)
  Set Interface =Nothing
End Sub

Public Sub ENVIODOCUMENTO_OnClick()
	Dim vDllDocumento As Object

	Set vDllDocumento = CreateBennerObject("BennerSaudeDesktopAtendimentoDocumentos.Rotinas")

	vDllDocumento.Executar(CurrentSystem)

	Set vDllDocumento = Nothing
End Sub

Public Sub FECHAAUTORIZACAO_OnClick()
    Dim Interface As Object
    Set Interface =CreateBennerObject("SamAuto.Autorizador")
    Interface.FecharSAM_AUTORIZ(CurrentSystem)
    Set Interface =Nothing
    'MsgBox "Processo de fechamento em manutenção"
End Sub



Public Sub FINALIZARPENDENTES_OnClick()
Dim UPD As Object
Dim SQL As Object

Set UPD =NewQuery
Set SQL =NewQuery


SQL.Add("SELECT PRAZOFINALIZARATENDIMENTO FROM SAM_PARAMETROSATENDIMENTO")
SQL.Active =True
If SQL.FieldByName("PRAZOFINALIZARATENDIMENTO").IsNull Then
MsgBox "Parâmentro prazo para finalizar não está definido - verificar parametros de atendimento"
Else
Dim DataPrazo As Date
DataPrazo =DateAdd("h",-SQL.FieldByName("PRAZOFINALIZARATENDIMENTO").AsInteger,ServerNow)

UPD.Add("UPDATE CA_ATEND")
UPD.Add("   SET DATAHORAFINAL = (SELECT MAX(A.DATAHORAINICIO) FROM CA_ATEND_LOG A WHERE A.ATENDIMENTO = CA_ATEND.HANDLE),")
UPD.Add("       FINALIZADOPELOSISTEMA = 'S'")
UPD.Add(" WHERE (DATAHORAINICIAL < :DATAPRAZO)")
UPD.Add("   AND DATAHORAFINAL IS NULL")
UPD.ParamByName("DATAPRAZO").Value =DataPrazo
'UPD.ParamByName("DATAATUAL").Value=ServerDate
UPD.ExecSQL
UPD.Clear
UPD.Add("UPDATE CA_ATEND")
UPD.Add("   SET DATAHORAFINAL = DATAHORAINICIAL,")
UPD.Add("       FINALIZADOPELOSISTEMA = 'S'")
UPD.Add(" WHERE (DATAHORAINICIAL < :DATAPRAZO)")
UPD.Add("   AND DATAHORAFINAL IS NULL")
UPD.ParamByName("DATAPRAZO").Value =DataPrazo
'UPD.ParamByName("DATAATUAL").Value=ServerDate
UPD.ExecSQL
MsgBox "Os registros foram atualizados!"
End If
Set UPD =Nothing
Set SQL =Nothing
End Sub

Public Sub MONITORATENDIMENTO_OnClick()
    Dim Interface As Object
    Set Interface =CreateBennerObject("CA028.MonitorUsuario")
    Interface.Exec(CurrentSystem)
    Set Interface =Nothing
End Sub

Public Sub MONITORAUDITPERICIA_OnClick()
      Dim Interface As Object
      Set Interface =CreateBennerObject("ca003.monitoraudit")
      Interface.Exec(CurrentSystem,0)
      Set Interface =Nothing
End Sub

Public Sub MONITORAUTORIZTRANSF_OnClick()
      Dim Interface As Object
      Set Interface =CreateBennerObject("ca004.MonitorAutorizTransf")
      Interface.Executar(CurrentSystem,0)
      Set Interface =Nothing
End Sub

Public Sub MONITORFAX_OnClick()
    Dim Interface As Object
	Dim SQL As Object

	Set SQL = NewQuery

	SQL.Active = False
	SQL.Add("SELECT FLAGINTERFACEFAX FROM SAM_PARAMETROSATENDIMENTO ")
	SQL.Active = True

	If SQL.FieldByName("FLAGINTERFACEFAX").AsString = "S" Then
		Set Interface =CreateBennerObject("ca053.MonitorDoc")
		Interface.Executar(CurrentSystem, 0)
	Else
		Set Interface =CreateBennerObject("ca002.monitordoc")
		Interface.Exec(CurrentSystem,0)
	End If

	Set SQL = Nothing
	Set Interface =Nothing
End Sub

Public Sub MONITOROUTROSERVICOS_OnClick()
Dim Interface As Object
Set Interface =CreateBennerObject("CA021.Monitor")
Interface.Executar(CurrentSystem)
Set Interface =Nothing
End Sub

Public Sub MONITORRESPOSTAS_OnClick()
      Dim Interface As Object
      Set Interface =CreateBennerObject("ca018.MonitorRespAutoriz")
      Interface.Exec(CurrentSystem,0)
      Set Interface =Nothing
End Sub

Public Sub MONITORULTSOLICIT_OnClick()
	  Dim Interface As Object
      Set Interface =CreateBennerObject("CA007.Consultas")
      Interface.Exec(CurrentSystem,0,0)
      Set Interface =Nothing
End Sub

'SMS 63801 - Débora Rebello - 26/01/2007
Public Sub ODONTOGRAMA_OnClick()

  Dim Obj As Object
  Dim Interface As Object
  Dim HandleBeneficiario As Long

  Set Obj = CreateBennerObject("Procura.Procurar")

  HandleBeneficiario = Obj.Exec(CurrentSystem,"SAM_BENEFICIARIO","NOME|BENEFICIARIO",2,"Nome|Beneficiário","","Beneficiários",False,"","")

  If (HandleBeneficiario > 0) Then

	Set Interface = CreateBennerObject("BSCLI006.ROTINAS")
	Interface.Odontograma(CurrentSystem, 0, 0, 0, 0, 0, HandleBeneficiario)
	Set Interface = Nothing
  End If

  Set Obj =Nothing

End Sub


Public Sub MENSAGEM_OnClick()
  Dim voInterface As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String
  Dim vvContainer As CSDContainer

  Set voInterface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
  viRetorno = voInterface.Exec(CurrentSystem, _
                               1, _
                               "TV_SAM_MENSAGEM_USUARIO", _
                               "Mensagem", _
                               0, _
                               600, _
                               500, _
                               False, _
                               vsMensagem, _
                               vvContainer)


  Set voInterface =Nothing

End Sub
