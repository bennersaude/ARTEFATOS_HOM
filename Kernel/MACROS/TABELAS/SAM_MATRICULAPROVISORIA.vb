'HASH: C3D3D400065EAE3940BC939BE2BBFD87
'Macro: SAM_MATRICULAPROVISORIA
'#Uses "*bsShowMessage"

'Última alteração: Milton/17/01/2002 -SMS 5976

Option Explicit
Dim EstadodaTabela As Long

Public Sub BOTAOPROCESSAR_OnClick()
  Dim vHandle As Long
  Dim INSMATR As Object
  Dim SQL As Object
  Set SQL = NewQuery

  If Not VerUsuario Then
    Exit Sub
  End If

  Set INSMATR = NewQuery


  INSMATR.Add("INSERT INTO SAM_MATRICULA " + _
              "( HANDLE, MATRICULA, NOME, Z_NOME, INICIAIS, NOMEMAE, NOMEPAI, SEXO, CPF, RG, ORGAOEMISSOR, DATANASCIMENTO, DATAINGRESSO,  FILIAL) VALUES " + _
              "(:HANDLE,:MATRICULA,:NOME,:Z_NOME,:INICIAIS,:NOMEMAE,:NOMEPAI,:SEXO,:CPF,:RG,:ORGAOEMISSOR,:DATANASCIMENTO,:DATAINGRESSO, :FILIAL) ")

  If CurrentQuery.State = 1 And CurrentQuery.FieldByName("REGISTROPROCESSADO").AsString = "N" Then 'Não está em inserção
    If CurrentQuery.FieldByName("ACAO").AsString = "I" Then
      If Not InTransaction Then StartTransaction

      'SQL.Add("SELECT FILIALPADRAO FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HANDLE")
      'SQL.ParamByName("HANDLE").Value =CurrentUser
      'SQL.Active =True
      Dim prFilial As Long
      Dim prFilialProcessamento As Long
      Dim prMsg As String

      If BuscarFiliais(CurrentSystem, prFilial, prFilialProcessamento, prMsg)Then
        bsShowMessage("Problemas para selecionar Filial. - Processo abortado!!", "I")
        Exit Sub
      End If

      INSMATR.ParamByName("FILIAL").Value = prFilial


      'Set SQL =Nothing

      vHandle = NewHandle("SAM_MATRICULA")
      INSMATR.ParamByName("HANDLE").Value = vHandle
      INSMATR.ParamByName("MATRICULA").Value = vHandle
      INSMATR.ParamByName("NOME").Value = CurrentQuery.FieldByName("NOME").AsString
      INSMATR.ParamByName("Z_NOME").Value = TiraAcento(CurrentQuery.FieldByName("NOME").AsString, True)
      INSMATR.ParamByName("INICIAIS").Value = CurrentQuery.FieldByName("INICIAIS").AsString
      INSMATR.ParamByName("NOMEMAE").Value = CurrentQuery.FieldByName("NOMEMAE").AsString
      INSMATR.ParamByName("NOMEPAI").Value = CurrentQuery.FieldByName("NOMEPAI").AsString
      INSMATR.ParamByName("SEXO").Value = CurrentQuery.FieldByName("SEXO").AsString
      INSMATR.ParamByName("CPF").Value = CurrentQuery.FieldByName("CPF").AsString
      INSMATR.ParamByName("RG").Value = CurrentQuery.FieldByName("RG").AsString
      INSMATR.ParamByName("ORGAOEMISSOR").Value = CurrentQuery.FieldByName("ORGAOEMISSOR").AsString
      INSMATR.ParamByName("DATANASCIMENTO").Value = CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime
      INSMATR.ParamByName("DATAINGRESSO").Value = ServerDate
      INSMATR.ExecSQL
      If InTransaction Then Commit
    End If
    CurrentQuery.Edit
    CurrentQuery.FieldByName("REGISTROPROCESSADO").Value = "S"
    CurrentQuery.Post
    RefreshNodesWithTable("SAM_MATRICULAPROVISORIA")
  End If

  Set INSMATR = Nothing
End Sub

Public Sub NOME_OnExit()
  If Not CurrentQuery.FieldByName("NOME").IsNull Then
    NOMEMAE.SetFocus
  End If
End Sub

Public Sub TABLE_AfterInsert()
  If Not PodeAlterarIncluir Then
    RefreshNodesWithTable "SAM_MATRICULAPROVISORIA"
    CurrentQuery.Cancel
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  CanContinue = PodeAlterarIncluir
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  CanContinue = PodeAlterarIncluir
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim GeraIniciais As Object
  Dim GeraHomonimos As Object
  Dim Lista As Object
  Dim vNome, vLista As String
  Dim Resposta As Long
  Dim vRG As String
  Dim i As Long
  Dim x As String


  EstadodaTabela = CurrentQuery.State

  If CurrentQuery.FieldByName("REGISTROPROCESSADO").Value = "N" Then
    If Not CurrentQuery.FieldByName("CPF").IsNull Then
      If(Not IsValidCPF(CurrentQuery.FieldByName("CPF").Value))Or _
         ((CurrentQuery.FieldByName("CPF").AsString = "11111111111")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "22222222222")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "33333333333")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "44444444444")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "55555555555")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "66666666666")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "77777777777")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "88888888888")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "99999999999")Or _
         (CurrentQuery.FieldByName("CPF").AsString = "00000000000"))Then
      bsShowMessage("C.P.F. Inválido", "E")
      CPF.SetFocus
      CanContinue = False
    End If
  End If
  If Not(CurrentQuery.FieldByName("RG").IsNull)Then
    vRG = Trim(CurrentQuery.FieldByName("RG").AsString)
    For i = 1 To Len(vRG)
      x = Mid(vRG, i, 1)
      If x <"0" Or x >"9" Then
        i = 1000
        Exit For
      End If
    Next
    If i = 1000 Then
      bsShowMessage("R.G. Inválido", "E")
      RG.SetFocus
      CanContinue = False
    End If
    CurrentQuery.FieldByName("RG").Value = vRG
  End If
  Set GeraIniciais = CreateBennerObject("Matricula.Iniciais")
  Set GeraHomonimos = CreateBennerObject("Matricula.Geracao")
  Set Lista = NewQuery
  Lista.Add("SELECT LISTA FROM SAM_PARAMETROSBENEFICIARIO")
  Lista.Active = True
  vLista = Lista.FieldByName("LISTA").AsString
  vNome = TiraAcento(CurrentQuery.FieldByName("NOME").AsString, True)
  CurrentQuery.FieldByName("INICIAIS").Value = GeraIniciais.Exec(CurrentSystem, vNome, vLista)
  ExcluiRegistros
  Set Lista = Nothing
  Set GeraIniciais = Nothing
End If

'------------------Durval 07/11/2001-----------------------------------------------------------------
If CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime >ServerDate Then
  bsShowMessage("Data nascimento maior que data atual !", "E")
  CanContinue = False
End If
'----------------------------------------------------------------------------------------------------

If Len(CurrentQuery.FieldByName("NOME").AsString)<3 Then
  bsShowMessage("Informe um nome com mais de três caracteres.", "E")
  CanContinue = False
  Exit Sub
End If

End Sub

Public Sub TABLE_AfterPost
  Dim GeraHomonimos As Object
  Dim SQL As Object
  Dim vPossuiHomonimo, vAcao As String

  If CurrentQuery.FieldByName("REGISTROPROCESSADO").Value = "N" Then
    Set SQL = NewQuery
    Set GeraHomonimos = CreateBennerObject("Matricula.Geracao")
    SQL.Add("UPDATE SAM_MATRICULAPROVISORIA SET POSSUIHOMONIMO = :H, ACAO = :A WHERE HANDLE = :HANDLE")
    vPossuiHomonimo = GeraHomonimos.Homonimos(CurrentSystem, _
                      CurrentQuery.FieldByName("LOTE").AsInteger, _
                      CurrentQuery.FieldByName("MATRICULA").AsInteger, _
                      CurrentQuery.FieldByName("INICIAIS").AsString, _
                      CurrentQuery.FieldByName("SEXO").AsString, _
                      CurrentQuery.FieldByName("CPF").AsString, _
                      CurrentQuery.FieldByName("RG").AsString, _
                      CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime)
    'FormatDateTime2("yyyy-mm-dd-hh.mm.ss",CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime))

    If EstadodaTabela = 3 Then
      vAcao = IIf(vPossuiHomonimo = "S", "N", "I")
      SQL.Clear
      SQL.Add("UPDATE SAM_MATRICULAPROVISORIA SET POSSUIHOMONIMO = :H, ACAO = :A WHERE HANDLE = :HANDLE")
      SQL.ParamByName("H").Value = vPossuiHomonimo
      SQL.ParamByName("A").Value = vAcao
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
      SQL.ExecSQL
    Else
      SQL.Clear
      SQL.Add("UPDATE SAM_MATRICULAPROVISORIA SET POSSUIHOMONIMO = :H WHERE HANDLE = :HANDLE")
      SQL.ParamByName("H").Value = vPossuiHomonimo
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
      SQL.ExecSQL
    End If

    Set SQL = Nothing
    Set GeraHomonimos = Nothing
  End If
End Sub

Public Sub ExcluiRegistros
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("DELETE FROM SAM_MATRICULAHOMONIMA WHERE MATRICULAPROVISORIA=:PROVISORIA")
  SQL.ParamByName("PROVISORIA").Value = CurrentQuery.FieldByName("MATRICULA").Value

  SQL.ExecSQL
  Set SQL = Nothing
End Sub


Public Function PodeAlterarIncluir As Boolean
  Dim SQL As Object
  Dim Frase, vHandle As String
  Set SQL = NewQuery
  vHandle = Str(RecordHandleOfTable("SAM_MATRICULALOTE"))'Pega o Handle do Lote
  SQL.Add(GenSql("SAM_MATRICULALOTE", "LOTEPROCESSADO", "", "HANDLE  = " + vHandle))'Monta SQL
  SQL.Active = True 'Executa SQL
  PodeAlterarIncluir = True
  If SQL.FieldByName("LOTEPROCESSADO").Value = "S" Then 'Verifica se o Lote já foi processado
    MsgBox("Insclusão / Alteração não Permitida!!! Lote já Processado", , "Crítica da Matrícula")'Mensagem
    'CurrentQuery.Cancel                                                                   		'Aborta a inclusão
    'RefreshNodesWithTable "SAM_MATRICULAPROVISORIA"                                       		'Remonta o Three View
    PodeAlterarIncluir = False
  End If

  If PodeAlterarIncluir = True Then
    PodeAlterarIncluir = VerUsuario
  End If

  Set SQL = Nothing
End Function

Public Function VerUsuario As Boolean
  Dim LOTE As Object
  Set LOTE = NewQuery
  LOTE.Add("SELECT USUARIO FROM SAM_MATRICULALOTE A WHERE A.HANDLE = :HANDLE")
  LOTE.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("LOTE").AsInteger
  LOTE.Active = True
  If LOTE.FieldByName("USUARIO").AsInteger <>CurrentUser Then
    VerUsuario = False
    bsShowMessage("Operação cancelada. Usuário não é o responsável", "E")
  Else
    VerUsuario = True
  End If
  LOTE.Active = False
  Set LOTE = Nothing
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOPROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	End If
End Sub
