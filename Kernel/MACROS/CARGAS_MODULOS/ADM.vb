'HASH: E69BABA306C7D4E620CAAC681D95CBB2
Public Sub ADMINISTRACAO_OnClick()
Dim obj As Object
    Set obj =CreateBennerObject("CorpAdmin.MainControl")
        obj.Exec(CurrentSystem)
        Set obj =Nothing
End Sub
Public Sub CADASTRARTISSDEPARA_OnClick()
	Dim dePara As Object
	Set dePara = CreateBennerObject("TISSDEPARA.Rotinas")
	dePara.Cadastrar(CurrentSystem)

	Set dePara = Nothing
End Sub

Public Sub EXPORTARAJUDA_OnClick()
	Dim state As Long

	state = CurrentQuery.State

	MsgBox(CStr(state))

	Exit Sub


	Dim obj As Object
	Dim Erro As String

	Set obj = CreateBennerObject("DLL.Atividades")
	Erro = obj.CallCronos(CurrentSystem)

	Exit Sub

  Dim CSHelp As Object
  Set CSHelp =CreateBennerObject("CSHelp.Importador")
  CSHelp.Exec(False)
  Set CSHelp =Nothing
End Sub
Public Sub EXPORTAVERSOES_OnClick()
  Dim dllVerificaVersoes As Object
  Set dllVerificaVersoes =CreateBennerObject("BsVerificaVersoes.Rotinas")
  dllVerificaVersoes.ExportaVersoes(CurrentSystem)
  Set dllVerificaVersoes =Nothing
End Sub

Public Sub IMPORTARAJUDA_OnClick()

	
  Dim CSHelp As Object
  Set CSHelp =CreateBennerObject("CSHelp.Importador")
  CSHelp.Exec(True)
  Set CSHelp =Nothing
End Sub
Public Sub BAIXAS_OnClick()
Dim obj As Object
Set obj =CreateBennerObject("Importador.Baixas")
    obj.Exec(CurrentSystem)
    Set obj =Nothing
End Sub
Public Sub CONTABILIDADE_OnClick()
Dim obj As Object
    Set obj =CreateBennerObject("Importador.Contabilidade")
        obj.Exec(CurrentSystem)
        Set obj =Nothing
End Sub
Public Sub DOCUMENTOS_OnClick()
Dim obj As Object
    Set obj =CreateBennerObject("Importador.Documento")
        obj.Exec(CurrentSystem)
        Set obj =Nothing
End Sub

Public Sub SEGURANCA_OnClick()
Dim obj As Object
        Set obj =CreateBennerObject("CS.Security")
        obj.Exec(CurrentSystem)
        Set obj =Nothing
End Sub
Public Sub INICIARCONTINGENCIA_OnClick()

  Dim vContingencia As String
  Dim sql As Object
  Set sql =NewQuery
  sql.Clear
  sql.Add("SELECT EMCONTINGENCIA")
  sql.Add("  FROM Z_SISTEMA")
  sql.Active =True
  If sql.FieldByName("EMCONTINGENCIA").AsString ="S" Then
    If MsgBox("A base de dados se encontra em modo de contigência, deseja finalizar modo de contingência?",vbYesNo +vbDefaultButton2)=vbYes Then
      vContingencia = "N"
    Else
      Exit Sub
    End If
  Else
    If MsgBox("Será iniciado o processo de contingência para as autorizações, deseja continuar?",vbYesNo +vbDefaultButton2)=vbYes Then
      vContingencia = "S"
    Else
      Exit Sub
    End If
  End If

  If Not InTransaction Then StartTransaction

  Set sql =NewQuery
  sql.Add("UPDATE Z_SISTEMA SET EMCONTINGENCIA =:CONT , USUARIOCONTINGENCIA=:USU, DATAHORACONTINGENCIA=:DATA")
  sql.ParamByName("CONT").Value =vContingencia
  sql.ParamByName("USU").Value =CurrentUser
  sql.ParamByName("DATA").Value =ServerNow
  sql.ExecSQL
  sql.Clear

  If InTransaction Then Commit
End Sub

Public Sub MONITORIMPXMLLOTE_OnClick()
	Dim obj As Object
	Set obj = CreateBennerObject("Benner.Saude.Desktop.MonitorImportacaoXmlLote.Rotinas")
  	obj.Exec(CurrentSystem)
  	Set obj = Nothing
End Sub

Public Sub SINCRONIZARCONTINGEN_OnClick()

  Dim SINCRO As Object
  Set SINCRO =NewQuery
  SINCRO.Clear
  SINCRO.Add("SELECT EMCONTINGENCIA")
  SINCRO.Add("  FROM Z_SISTEMA")
  SINCRO.Active =True
  If SINCRO.FieldByName("EMCONTINGENCIA").AsString ="S" Then
    MsgBox("A base de dados se encontra em modo de contigência. O processo de sincronismo não pode ser executado!")
    Exit Sub
  Else
    If MsgBox("Será iniciado o processo de sincronismo das autorizações, deseja continuar?",vbYesNo +vbDefaultButton2)=vbNo Then
      Exit Sub
    End If
  End If

  Dim obj As Object
  Set obj =CreateBennerObject("CA052.Rotinas")
  obj.SincronizarContingencia(CurrentSystem)
  Set obj =Nothing
End Sub

Public Sub UNIFICARREGISTROS_OnClick()
Dim obj As Object
        Set obj =CreateBennerObject("CS.UnifyForm")
        obj.Exec()
        Set obj =Nothing
End Sub

Public Sub VERIFICAVERSOES_OnClick()
  Dim dllVerificaVersoes As Object
  Set dllVerificaVersoes =CreateBennerObject("BsVerificaVersoes.Rotinas")
  dllVerificaVersoes.VerificaVersoes(CurrentSystem)
  Set dllVerificaVersoes =Nothing
End Sub

Public Sub MODULE_BeforeNodeShow(ByVal NodeFullPath As String, CanShow As Boolean)
	If Right(NodeFullPath, 29) = "|SAM_REGRASAPROVACAO_USUARIOS" Or Right(NodeFullPath, 29) = "|SAM_REGRASAPROVACAO_GRUPOSEG" Then

		Dim vQueryPermissao As BPesquisa
		Set vQueryPermissao = NewQuery
		vQueryPermissao.Add("SELECT GRUPOSEGURANCA, USUARIOSSELECIONADOS")
		vQueryPermissao.Add("        FROM SAM_REGRASAPROVACAO           ")
        vQueryPermissao.Add("WHERE HANDLE = :PHANDLE                    ")

		vQueryPermissao.ParamByName("PHANDLE").Value = RecordHandleOfTable("SAM_REGRASAPROVACAO")
		vQueryPermissao.Active = True

		If Right(NodeFullPath, 29) = "|SAM_REGRASAPROVACAO_USUARIOS" Then
			CanShow = vQueryPermissao.FieldByName("USUARIOSSELECIONADOS").AsString = "S"
		End If

		If Right(NodeFullPath, 29) = "|SAM_REGRASAPROVACAO_GRUPOSEG" Then
			CanShow = vQueryPermissao.FieldByName("GRUPOSEGURANCA").AsString = "S"
		End If

		Set vQueryPermissao = Nothing
	End If
End Sub
