'HASH: E7F0302C713F8FCC957E4CC835A36C4F
'#USES "*CriaTabelaTemporariaSqlServer"
Option Explicit
'Dim vDLLFinanceiro As Object
' BUILDER

Public Sub QUERY_OnClick()

  Dim interface As Object
  
  Set interface=CreateBennerObject("SamQuery.QueryBuilder")
  interface.Exec
  Set interface=Nothing
End Sub
Public Sub Application_OnClose(CanClose As Boolean)

  Dim interface1 As Object
  Set interface1=CreateBennerObject("BSDEBUG.ROTINAS")
  On Error GoTo pula1
  interface1.Sair
  If interface1.VerificaAberto=True Then
    MsgBox("Favor finalizar a tela de depuração")
    CanClose=False
    Set interface1=Nothing
    Exit Sub
  End If

  pula1: Set interface1=Nothing


  Dim interface As Object

  Set interface=CreateBennerObject("SAMPEG.PROCESSAR")
  interface.finalizar
  Set interface=Nothing

  If Not InTransaction Then
    StartTransaction
  End If
  'Augusto SMS 41171 - Item 2
  UserVar("BENEFICIARIO") =""

  Dim DEL As Object
  Set DEL = NewQuery
  DEL.Add("DELETE FROM CA_ULTIMOUSUARIO")
  DEL.Add(" WHERE ULTIMOUSUARIOREVERSAO = :USUARIO OR USUARIO = :USUARIO")
  DEL.ParamByName("USUARIO").Value = CurrentUser
  DEL.ExecSQL
  Set DEL = Nothing

  If InTransaction Then
    Commit
  End If

  'Problemas de memória na Cassi. É necessário incializar o financeiro.geral aqui - Rodrigo - sms 48887
  'vDLLFinanceiro.Finalizar
  'Set vDLLFinanceiro = Nothing

  
End Sub

Public Sub Application_OnOpen(CanContinue As Boolean)


  	'teste

  	'jogando vazio nas variaveis do leiaute da guia para carregar o leiaute padrão quando entrar
  	UserVar("CAMPOS_LEIAUTE_GUIA") = ""
  	UserVar("CAMPOS_LEIAUTE_GUIA_EVENTO") = ""
  	UserVar("BSDEBUG")="N"
  	UserVar("BSDEBUGDE")="N"
  	UserVar("interlocutor") = ""

	Dim I As Integer
	Dim T(100) As String
	Dim SQL As Object
	Dim DEL As Object


	Set DEL = NewQuery
	Set SQL = NewQuery

	'SQL.Add("SELECT NOME")
    'SQL.Add("  FROM Z_TABELAS")
    'SQL.Add(" WHERE NOME LIKE 'TMP%'")
	' SQL.Add("AND EXISTS (SELECT HANDLE FROM Z_CAMPOS WHERE TABELA = Z_TABELAS.HANDLE AND NOME = 'DATADECRIACAO')")

	'jogando vazio nas variaveis do leiaute da guia para carregar o leiaute padrão quando entrar
  	UserVar("CAMPOS_LEIAUTE_GUIA") = ""
  	UserVar("CAMPOS_LEIAUTE_GUIA_EVENTO") = ""
  	UserVar("BSDEBUG")="N"
  	UserVar("BSDEBUGDE")="N"

	'I=1
	'SQL.Active=True
	'While Not SQL.EOF
	'	T(I)=SQL.FieldByName("NOME").AsString
	'	SQL.Next
	'	I=I+1
	'Wend

	'I=1
	'While I<100
	'	DEL.Clear
	'	If T(I)<>""	Then
	'		DEL.Add("DELETE FROM "+T(I) +" WHERE DATADECRIACAO < :DATA")
	'		DEL.ParamByName("DATA").Value=(ServerDate() - 5)
	'		DEL.ExecSQL
	'	End If
	'	I = I+1
	'Wend



	


    'SMS 45152 Wagner Santos 28/07/2005 - Grava ultimo acesso do usuário.
    If Not InTransaction Then
      StartTransaction
    End If
    SQL.Add("UPDATE Z_GRUPOUSUARIOS")
    SQL.Add(" SET DATAHORAULTIMOACESSO = :DATA")
    SQL.Add("WHERE HANDLE = :USUARIO")
    SQL.ParamByName("USUARIO").Value = CurrentUser
    SQL.ParamByName("DATA").Value = ServerNow()
    SQL.ExecSQL
    'Fim SMS 45152

 	DEL.Add("DELETE FROM CA_ULTIMOUSUARIO")
  	DEL.Add(" WHERE ULTIMOUSUARIOREVERSAO = :USUARIO OR USUARIO = :USUARIO")
  	DEL.ParamByName("USUARIO").Value = CurrentUser
  	DEL.ExecSQL
    If InTransaction Then
      Commit
    End If

	Set SQL = Nothing
	Set DEL = Nothing

	'Problemas de memória na Cassi. É necessário incializar o financeiro.geral aqui - Rodrigo - sms 48887
	'Set vDLLFinanceiro = CreateBennerObject("FINANCEIRO.GERAL")
    'vDLLFinanceiro.Inicializar(CurrentSystem)

    'SMS 68254 - Marcelo Barbosa - 04/10/2006
    If Not InTransaction Then
      StartTransaction
    End If

    If InStr(SQLServer, "MSSQL")>0 Then
        CriaTabelaTemporariaSqlServer
    ElseIf InStr(SQLServer, "DB2")>0 Then    'SMS 84247 - Ricardo Vieira - 21/08/2007
		CreateTemporaryTable("TMP_LIMITE")
		CreateTemporaryTable("TMP_NEGACAOEVENTO")
		CreateTemporaryTable("TMP_ALERTAS")
		CreateTemporaryTable("TMP_MENSAGEM")
		CreateTemporaryTable("TMP_ORIGEMCALCULO")
		CreateTemporaryTable("TMP_PRAZOQUANT")
		CreateTemporaryTable("TMP_OBSERVACAO")
		CreateTemporaryTable("TMP_INCOMPATIBILIDADES")
		CreateTemporaryTable("TMP_QUANTIDADES_PF")
		CreateTemporaryTable("TMP_EVENTOGERADO")
		CreateTemporaryTable("TMP_REDERESTRITA")
		CreateTemporaryTable("TMP_PRESTADOR_CONSULTA")
		CreateTemporaryTable("TMP_CONSULTA")
	    CreateTemporaryTable("TMP_ENDERECO") 'SMS 91130 - Willian
		CreateTemporaryTable("TMP_CANCELAGUIA")
		CreateTemporaryTable("TMP_AUX1")
		CreateTemporaryTable("TMP_PRECOPACOTE")
		CreateTemporaryTable("TMP_MSGREGULARIZAREVENTO") 'SMS 100227 - Paulo Melo
		CreateTemporaryTable("TMP_VALORFXEVENTO") 'SMS 101179 - Rafael Canali

    End If

    If InTransaction Then
      Commit
    End If

End Sub

Public Sub AUDITORIA_OnClick()
  If CurrentTable>0 Then  
     Dim registro As Long
     On Error GoTo def
       registro=CurrentQuery.FieldByName("HANDLE").AsInteger
       
       
       GoTo faz
     def:
       registro=0
     
     faz:
     Dim interface As Object
     Set interface=CreateBennerObject("samutil.auditoria")
     interface.Exec(CurrentSystem, CurrentTable,registro)
     Set interface=Nothing    
     
  End If

End Sub

Public Sub DEPURAR_OnClick()
	Dim SQL As Object
	Set SQL = NewQuery
	SQL.Add("SELECT G.BOTOES FROM Z_GRUPOS G, Z_GRUPOUSUARIOS U WHERE G.HANDLE = U.GRUPO AND U.HANDLE = :HANDLE")
	SQL.ParamByName("HANDLE").Value = CurrentUser
	SQL.Active = True
	If SQL.FieldByName("BOTOES").AsInteger >= 512 Then
	    Dim interface As Object
	    Set interface=CreateBennerObject("BSDebug.Rotinas")
	    interface.Show(CurrentSystem)
	    Set interface=Nothing
	Else
		MsgBox "Usuário não tem permissão para utilizar a depuração."
	End If
End Sub

Public Sub MONITOR_OnClick()
	Dim interface As Object
    Set interface=CreateBennerObject("BSMONITOR.PROCESSO")
    interface.Executar(CurrentSystem)
    Set interface=Nothing
End Sub

Public Sub MONITOROUTSERVICOS_OnClick()
	Dim qAux As Object
	Dim interface As Object
	Set qAux = NewQuery
 	qAux.Add("SELECT MONITOROUTSERVCOMANDOSAPL FROM SAM_PARAMETROSATENDIMENTO")
 	qAux.Active = True
 	If qAux.FieldByName("MONITOROUTSERVCOMANDOSAPL").AsString = "S" Then
 	   qAux.Active = False
       Set qAux = Nothing
       Set interface = CreateBennerObject("CA021.MONITOR")
	   interface.Executar(CurrentSystem)
	   Set interface = Nothing
	Else
 	   qAux.Active = False
 	   Set qAux = Nothing
 	   MsgBox("Não é possível acessar o Monitor de Outros Serviços por meio dos Comandos da Aplicação")
 	End If

End Sub
