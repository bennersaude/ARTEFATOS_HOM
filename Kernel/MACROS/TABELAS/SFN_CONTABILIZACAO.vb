'HASH: AFA1680E8B5AA4DE0FD140E0644A8287
'Macro: SFN_CONTABILIZACAO
'#Uses "*bsShowMessage"

Option Explicit
Dim vOPERACAO As Integer
Dim vTABCLASSEDEB As Integer
Dim vCLASSECONTABILDEB As Integer
Dim vTABCLASSECRE As Integer
Dim vCLASSECONTABILCRE As Integer
Dim vCONTABHIST As Integer

Dim vConcluido As Boolean
Dim vSituacao As Integer

Public Sub BOTAOCAMPOS_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("Financeiro.Campos")
  interface.Exec(CurrentSystem)
  Set interface = Nothing

End Sub

Public Sub CONTABHIST_OnPopup(ShowPopup As Boolean)
  Dim interface As Object

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_CONTABHIST.CODIGO|SFN_CONTABHIST.HISTORICO"

  vCriterio = ""

  vCampos = "Código|Histórico"

  vHandle = interface.Exec(CurrentSystem, "SFN_CONTABHIST", vColunas, 1, vCampos, vCriterio, "Histórico Padrão", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTABHIST").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub TABLE_AfterInsert()
  vSituacao = CurrentQuery.State
End Sub

Public Sub TABLE_AfterPost()
  If vSituacao = 3 Then
    'RotinaInclusao
    vConcluido = True
  'Else
  '  RotinaAlteracao
  End If
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		CONTABHIST.WebLocalWhere = ""
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vOPERACAO = CurrentQuery.FieldByName("OPERACAO").AsInteger
  vTABCLASSEDEB = CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger
  vCLASSECONTABILDEB = CurrentQuery.FieldByName("CLASSECONTABILDEB").AsInteger
  vTABCLASSECRE = CurrentQuery.FieldByName("TABCLASSECRE").AsInteger
  vCLASSECONTABILCRE = CurrentQuery.FieldByName("CLASSECONTABILCRE").AsInteger
  vCONTABHIST = CurrentQuery.FieldByName("CONTABHIST").AsInteger

  If (vSituacao = 0) Or (vSituacao = 1 And vConcluido) Then vSituacao = CurrentQuery.State
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  Dim SQLPesq As Object
  Set SQLPesq = NewQuery

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT CODIGO FROM SIS_OPERACAO WHERE HANDLE=:HOPERACAO")
  SQL.ParamByName("HOPERACAO").AsInteger = CurrentQuery.FieldByName("OPERACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("CODIGO").AsInteger = 131 Then 'LUCRO/PERDA
    If(CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger <>1)Or(CurrentQuery.FieldByName("TABCLASSECRE").AsInteger <>1)Then
      bsShowMessage("Classificação 'Debita/Credita' somente podem ser 'FIXA' para operação: '131.Baixa referente lucro/perda'", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If

  If SQL.FieldByName("CODIGO").AsInteger = 110 Then 'LANÇAMENTO
    If(CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger = 3)Or(CurrentQuery.FieldByName("TABCLASSECRE").AsInteger = 3)Then
      bsShowMessage("Classificação 'Debita/Credita' NÃO pode ser 'CONTA TESOURARIA' para operação: '110.Lançamento Fatura'", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If

 ' Coelho SMS: 68853 - fazendo a validação do cadastramento de regra já existente
 SQLPesq.Clear
 SQLPesq.Active = False
 SQLPesq.Add("SELECT COUNT(HANDLE) QTDE   ")
 SQLPesq.Add("  FROM SFN_CONTABILIZACAO   ")
 SQLPesq.Add(" WHERE OPERACAO=:OP         ")
 SQLPesq.Add("   AND HANDLE<>:HG          ")
 SQLPesq.Add("   AND CLASSEGERENCIAL=:CG  ")
 SQLPesq.ParamByName("HG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 SQLPesq.ParamByName("OP").AsInteger = CurrentQuery.FieldByName("OPERACAO").AsInteger
 SQLPesq.ParamByName("CG").AsInteger = CurrentQuery.FieldByName("CLASSEGERENCIAL").AsInteger
 SQLPesq.Active = True

 If SQLPesq.FieldByName("QTDE").AsInteger >0 Then 'REGISTRO JA EXISTE
   bsShowMessage("Já existe Regra de Contabilização para este Tipo de Operação!!!", "E")
   CanContinue = False
   Exit Sub
 End If

Set SQL = Nothing
Set SQLPesq = Nothing

End Sub

Public Sub RotinaAlteracao
  Dim vClasseD As Integer
  Dim vClasseC As Integer
  Dim vHistor  As Integer
  Dim QueryUpdate As Object
  Dim Query As Object

  If vTABCLASSEDEB <> CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger Or _
                                                vCLASSECONTABILDEB <> CurrentQuery.FieldByName("CLASSECONTABILDEB").AsInteger Or _
                                                vTABCLASSECRE <> CurrentQuery.FieldByName("TABCLASSECRE").AsInteger Or _
                                                vCLASSECONTABILCRE <> CurrentQuery.FieldByName("CLASSECONTABILCRE").AsInteger Or _
                                                vCONTABHIST <> CurrentQuery.FieldByName("CONTABHIST").AsInteger Then

    If bsShowMessage("Deseja replicar a(s) alteração(ões) para outras classes gerenciais ?", "Q") = vbYes Then
      Set Query = NewQuery
      Set QueryUpdate = NewQuery
      Dim vNivel As String

      vNivel = InputBox("Estrutura", "Informe o nível de classe para replicação")

      If vNivel = "" Then
        bsShowMessage("O nível da classe para replicação não foi informado", "I")
        Exit Sub
      End If

      Query.Clear
      Query.Add("SELECT HANDLE                   ")
      Query.Add("  FROM SFN_CLASSEGERENCIAL      ")
      Query.Add(" WHERE ESTRUTURA = :ESTRUTURA   ")
      Query.ParamByName("ESTRUTURA").Value = vNivel
      Query.Active = True
      If Query.EOF Then
        bsShowMessage("O nível da classe para replicação não existe", "I")
        Exit Sub
      End If

      'Query.Clear
      'Query.Add("SELECT C.* FROM SFN_CLASSEGERENCIAL A")
      'Query.Add("  JOIN SFN_CONTABILIZACAO R ON (R.CLASSEGERENCIAL = A.HANDLE)")
      'Query.Add("  JOIN SFN_CLASSEGERENCIAL B ON (A.NIVELSUPERIOR = B.NIVELSUPERIOR AND B.ULTIMONIVEL = 'S')")
      'Query.Add("  JOIN SFN_CONTABILIZACAO C ON (C.CLASSEGERENCIAL = B.HANDLE)")
      'Query.Add(" WHERE A.HANDLE = :PCLASSEGERENCIAL")
      'Query.Add("   AND R.OPERACAO = C.OPERACAO")
      'Query.Add("   AND R.OPERACAO = :POPERACAO")
      'Query.ParamByName("POPERACAO").AsInteger = CurrentQuery.FieldByName("OPERACAO").AsInteger
      'Query.ParamByName("PCLASSEGERENCIAL").AsInteger = CurrentQuery.FieldByName("CLASSEGERENCIAL").AsInteger
      'Query.Active = True 
      Query.Clear
      Query.Add("SELECT C.*                                                  ")
      Query.Add("FROM SFN_CLASSEGERENCIAL A                                  ")
      Query.Add("JOIN SFN_CONTABILIZACAO C  ON (C.CLASSEGERENCIAL =A.HANDLE) ")
      vNivel = "'" + vNivel + "%'"
      Query.Add(" WHERE A.HANDLE <> :PCLASSEGERENCIALATUAL                   ")
      Query.Add("   AND A.ESTRUTURA LIKE " + vNivel)
      Query.Add("   AND A.ULTIMONIVEL = 'S'                                  ")
      Query.Add("   AND C.OPERACAO = :POPERACAO                              ")
      Query.ParamByName("POPERACAO").AsInteger = CurrentQuery.FieldByName("OPERACAO").AsInteger
      Query.ParamByName("PCLASSEGERENCIALATUAL").AsInteger = CurrentQuery.FieldByName("CLASSEGERENCIAL").AsInteger
      Query.Active = True

    '  If Not InTransaction Then StartTransaction


      While Not Query.EOF

        vClasseD = 0
        vClasseC = 0
        vHistor  = 0

        QueryUpdate.Clear
        QueryUpdate.Add("UPDATE SFN_CONTABILIZACAO                      ")
        QueryUpdate.Add("   SET TABCLASSEDEB =:PTABCLASSEDEB,           ")

        If vTABCLASSEDEB <> CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger Then
          QueryUpdate.ParamByName("PTABCLASSEDEB").AsInteger = CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger
        Else
          QueryUpdate.ParamByName("PTABCLASSEDEB").AsInteger = Query.FieldByName("TABCLASSEDEB").AsInteger
        End If

        If CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger = 1 Then

          If vCLASSECONTABILDEB <> CurrentQuery.FieldByName("CLASSECONTABILDEB").AsInteger Then
            vClasseD = CurrentQuery.FieldByName("CLASSECONTABILDEB").AsInteger
          Else
            vClasseD = Query.FieldByName("CLASSECONTABILDEB").AsInteger
          End If

          If vClasseD > 0 Then
            QueryUpdate.Add("       CLASSECONTABILDEB =:PCLASSECONTABILDEB, ")
            QueryUpdate.ParamByName("PCLASSECONTABILDEB").AsInteger = vClasseD
          End If

        End If

        If CurrentQuery.FieldByName("TABCLASSECRE").AsInteger = 1 Then

          If vCLASSECONTABILCRE <> CurrentQuery.FieldByName("CLASSECONTABILCRE").AsInteger Then
            vClasseC = CurrentQuery.FieldByName("CLASSECONTABILCRE").AsInteger
          Else
            vClasseC = Query.FieldByName("CLASSECONTABILCRE").AsInteger
          End If

          If vClasseC > 0 Then
            QueryUpdate.Add("       CLASSECONTABILCRE =:PCLASSECONTABILCRE, ")
            QueryUpdate.ParamByName("PCLASSECONTABILCRE").AsInteger = vClasseC
          End If

        End If


        If vCONTABHIST <> CurrentQuery.FieldByName("CONTABHIST").AsInteger Then
          vHistor = CurrentQuery.FieldByName("CONTABHIST").AsInteger
        Else
          vHistor = Query.FieldByName("CONTABHIST").AsInteger
        End If

        If vHistor > 0 Then
          QueryUpdate.Add("       CONTABHIST =:PCONTABHIST,               ")
          QueryUpdate.ParamByName("PCONTABHIST").AsInteger = vHistor
        End If

        QueryUpdate.Add("       TABCLASSECRE =:PTABCLASSECRE            ")
        If vTABCLASSECRE <> CurrentQuery.FieldByName("TABCLASSECRE").AsInteger Then
          QueryUpdate.ParamByName("PTABCLASSECRE").AsInteger = CurrentQuery.FieldByName("TABCLASSECRE").AsInteger
        Else
          QueryUpdate.ParamByName("PTABCLASSECRE").AsInteger = Query.FieldByName("TABCLASSECRE").AsInteger
        End If

        QueryUpdate.Add(" WHERE HANDLE = :PHANDLE                       ")

        QueryUpdate.ParamByName("PHANDLE").AsInteger = Query.FieldByName("HANDLE").AsInteger
        QueryUpdate.ExecSQL

        Query.Next
      Wend
'      If InTransaction Then Commit

      Set Query = NewQuery
      Set QueryUpdate = NewQuery
    End If
  End If


End Sub


Public Sub RotinaInclusao
	Dim Query As Object
	Dim QueryInsert As Object
  If bsShowMessage("Deseja efetivar esta inclusão, para as demais classes gerenciais ?", "Q") = vbYes Then
    Set Query = NewQuery
    Set QueryInsert = NewQuery

    Query.Clear
    Query.Add("SELECT B.* FROM SFN_CLASSEGERENCIAL A")
    Query.Add("  JOIN SFN_CLASSEGERENCIAL B ON (A.NIVELSUPERIOR = B.NIVELSUPERIOR AND B.ULTIMONIVEL = 'S')")  ' Paulo Melo - Invertido os joins para nao dar problema em Oracle.
    Query.Add("  JOIN SFN_CONTABILIZACAO R ON (R.CLASSEGERENCIAL = A.HANDLE)")
    Query.Add(" WHERE A.HANDLE = :PCLASSEGERENCIAL")
    Query.Add("   AND NOT EXISTS (SELECT HANDLE FROM SFN_CONTABILIZACAO C WHERE C.CLASSEGERENCIAL =B.HANDLE AND C.OPERACAO = R.OPERACAO)")
    Query.Add("   AND R.OPERACAO = :POPERACAO")
    Query.ParamByName("POPERACAO").AsInteger = CurrentQuery.FieldByName("OPERACAO").AsInteger
    Query.ParamByName("PCLASSEGERENCIAL").AsInteger = CurrentQuery.FieldByName("CLASSEGERENCIAL").AsInteger
    Query.Active = True

    QueryInsert.Clear

    QueryInsert.Add("INSERT INTO SFN_CONTABILIZACAO (HANDLE,CLASSEGERENCIAL,OPERACAO,TABCLASSEDEB, CLASSECONTABILDEB, TABCLASSECRE, CLASSECONTABILCRE, CONTABHIST)")
    QueryInsert.Add("VALUES (:PHANDLE,:PCLASSEGERENCIAL,:POPERACAO, :PTABCLASSEDEB, :PCLASSECONTABILDEB, :PTABCLASSECRE, :PCLASSECONTABILCRE, :PCONTABHIST)")


    'If Not InTransaction Then StartTransaction


    While Not Query.EOF
      QueryInsert.ParamByName("PHANDLE").AsInteger = NewHandle("SFN_CONTABILIZACAO")
      QueryInsert.ParamByName("POPERACAO").AsInteger = CurrentQuery.FieldByName("OPERACAO").AsInteger
      QueryInsert.ParamByName("PTABCLASSEDEB").AsInteger = CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger

      If CurrentQuery.FieldByName("TABCLASSEDEB").AsInteger = 1 Then
        QueryInsert.ParamByName("PCLASSECONTABILDEB").AsInteger = CurrentQuery.FieldByName("CLASSECONTABILDEB").AsInteger
      Else
        QueryInsert.ParamByName("PCLASSECONTABILDEB").DataType = ftInteger
        QueryInsert.ParamByName("PCLASSECONTABILDEB").Clear
      End If

      QueryInsert.ParamByName("PTABCLASSECRE").AsInteger = CurrentQuery.FieldByName("TABCLASSECRE").AsInteger

      If CurrentQuery.FieldByName("TABCLASSECRE").AsInteger = 1 Then
        QueryInsert.ParamByName("PCLASSECONTABILCRE").AsInteger = CurrentQuery.FieldByName("CLASSECONTABILCRE").AsInteger
      Else
        QueryInsert.ParamByName("PCLASSECONTABILCRE").DataType = ftInteger
        QueryInsert.ParamByName("PCLASSECONTABILCRE").Clear
      End If

      QueryInsert.ParamByName("PCONTABHIST").AsInteger = CurrentQuery.FieldByName("CONTABHIST").AsInteger


      QueryInsert.ParamByName("PCLASSEGERENCIAL").AsInteger = Query.FieldByName("HANDLE").AsInteger
      QueryInsert.ExecSQL

      Query.Next
    Wend
    'If InTransaction Then Commit

    Set Query = NewQuery
    Set QueryInsert = NewQuery
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOCAMPOS" Then
		BOTAOCAMPOS_OnClick
	End If
End Sub
