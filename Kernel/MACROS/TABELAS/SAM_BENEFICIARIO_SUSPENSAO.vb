'HASH: 8DEEE4B47EBED37F3F589E2181093056
'Macro: SAM_BENEFICIARIO_SUSPENSAO
'#Uses "*bsShowMessage"

Public Sub MOTIVOSUSPENSAO_OnPopup(ShowPopup As Boolean)
  Dim vsValor As String
  Dim q1 As Object
  Set q1 = NewQuery

  q1.Active = False
  q1.Clear
  q1.Add("Select B.CODIGO ")
  q1.Add("FROM SAM_CONTRATO A ")
  q1.Add("Join SIS_TIPOFATURAMENTO B On (B.HANDLE = A.TIPOFATURAMENTO)  ")
  q1.Add("WHERE A.HANDLE = :HANDLE  ")
  q1.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_CONTRATO")
  q1.Active = True

  vsValor = "3"

  If q1.FieldByName("CODIGO").AsInteger = 130 Then
    vsValor = "-1"
  End If

  Set q1 = Nothing

  MOTIVOSUSPENSAO.LocalWhere = "TABTIPO <> "+vsValor+""		' SMS 95929 - Paulo Melo - 23/04/2008 - Campos TAB devem ser comparados a INTEIROS, para não dar pau em DB2

End Sub

Public Sub TABLE_AfterInsert()

End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If

  If CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    COMPETENCIAFINAL.ReadOnly = False
  Else
    COMPETENCIAFINAL.ReadOnly = True
  End If
  If CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull Then
    COMPETENCIAINICIAL.ReadOnly = False
  Else
    COMPETENCIAINICIAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CanContinue = False
    bsShowMessage("Registro finalizado não pode ser alterado!", "E")
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not(CurrentQuery.FieldByName("DATAFINAL").IsNull)And _
         (CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)Then
    CanContinue = False
    bsShowMessage("A data final não pode ser inferior à data inicial", "E")
    Exit Sub
  End If

  If(Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And(CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull)Then
    CanContinue = False
    bsShowMessage("A competência final não pode ser preenchida quando a competência inicial for nula.", "E")
    Exit Sub
  End If

Dim q1 As Object
Set q1 = NewQuery

q1.Active = False
q1.Clear
q1.Add("SELECT SUSPENDEFATURAMENTO,TABTIPO  FROM SAM_MOTIVOSUSPENSAO WHERE HANDLE = :HANDLE")
q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MOTIVOSUSPENSAO").AsInteger
q1.Active = True

If ( (q1.FieldByName("SUSPENDEFATURAMENTO").AsString = "S")  Or (q1.FieldByName("TABTIPO").AsString = "3")  ) And (CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull)  Then
  bsShowMessage("Competência inicial é obrigatório.", "E")
  CanContinue = False
  Exit Sub
Else
  If (q1.FieldByName("SUSPENDEFATURAMENTO").AsString <>"S") And (Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull) And (q1.FieldByName("TABTIPO").AsString <> "3") Then
    bsShowMessage("Competência inicial deve ser nula.", "E")
    CanContinue = False
    Exit Sub
  End If
End If

If(CurrentQuery.FieldByName("FATURARMESESSUSPENSOS").AsString = "S")And(CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)Then
CanContinue = False
bsShowMessage("Não é possível faturar meses suspensos sem que a competência final esteja preenchida.", "E")
Exit Sub
End If

If Not(CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And _
       (CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime <CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)Then
  CanContinue = False
  bsShowMessage("A competência final não pode ser inferior à competência inicial", "E")
  Exit Sub
End If

Dim SQL As Object
Dim SQL1 As Object
Set SQL = NewQuery
Set SQL1 = NewQuery

SQL.Clear
SQL.Add("SELECT DATAADESAO")
SQL.Add("FROM SAM_BENEFICIARIO")
SQL.Add("WHERE HANDLE = :HBENEFICIARIO")
SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
SQL.Active = True
If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
  CanContinue = False
  Set SQL = Nothing
  Set SQL1 = Nothing
  bsShowMessage("A data inicial não pode ser inferior à adesão do beneficiário", "E")
  Exit Sub
End If

Dim vMesComp As Integer
Dim vAnoComp As Integer
Dim vMesAdesao As Integer
Dim vAnoAdesao As Integer

vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)
vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)
vMesAdesao = DatePart("m", SQL.FieldByName("DATAADESAO").AsDateTime)
vAnoAdesao = DatePart("yyyy", SQL.FieldByName("DATAADESAO").AsDateTime)

If Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull Then
  If(vAnoComp <vAnoAdesao)Or _
     (vAnoComp = vAnoAdesao And vMesComp <vMesAdesao)Then
  CanContinue = False
  Set SQL = Nothing
  Set SQL1 = Nothing
  bsShowMessage("A competência inicial não pode ser inferior à adesão do beneficiário", "E")
  Exit Sub
End If
End If

Set SQL = Nothing
Set SQL1 = Nothing

If CurrentQuery.State = 3 Then
  Dim Interface As Object
  Dim Linha As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_SUSPENSAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
  Set Interface = Nothing
End If

CanContinue = CheckVigenciaBenef

'If CHECARCOMPSUSPENSAO Then
'CanContinue =False
'RefreshNodesWithTable("SAM_BENEFICIARIO_SUSPENSAO")
'Exit Sub
'End If

Dim SQLFechamento
Set SQLFechamento = NewQuery
SQLFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
SQLFechamento.Active = True

If CurrentQuery.State = 3 Then
  If SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime >CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
    bsShowMessage("Não é possível cadastrar data inicial inferior a data de fechamento - Parâmetros Gerais", "E")
    CanContinue = False
  End If
End If

If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
  If SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime >CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    bsShowMessage("Não é possível cadastrar data final inferior a data de fechamento - Parâmetros Gerais", "E")
    CanContinue = False
  End If
End If


Dim vMesComp1 As Integer
Dim vAnoComp1 As Integer
Dim vMesComp2 As Integer
Dim vAnoComp2 As Integer
Dim vMesFechamento As Integer
Dim vAnoFechamento As Integer

If Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull Then

  vMesComp1 = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)
  vAnoComp1 = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)
  vMesComp2 = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)
  vAnoComp2 = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)

  vMesFechamento = DatePart("m", SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
  vAnoFechamento = DatePart("yyyy", SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

  If CurrentQuery.State = 3 Then
    If(vAnoComp1 <vAnoAdesao)Or _
       (vAnoComp1 = vAnoAdesao And vMesComp1 <vMesFechamento)Then
    CanContinue = False
    bsShowMessage("A competência inicial não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
  End If
End If
End If

If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
  If(vAnoComp2 <vAnoAdesao)Or _
     (vAnoComp2 = vAnoAdesao And vMesComp2 <vMesFechamento)Then
  CanContinue = False
  bsShowMessage("A competência final não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
End If
End If


Set SQLFechamento = Nothing

End Sub

Public Function CheckVigenciaBenef As Boolean
  CheckVigenciaBenef = True
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DATAADESAO,DATACANCELAMENTO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEF")
  SQL.ParamByName("BENEF").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Data Inicial da Suspensão inferior a Adesão do Beneficiário!", "I")
    CheckVigenciaBenef = False
  Else
    If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >SQL.FieldByName("DATACANCELAMENTO").AsDateTime Then
        bsShowMessage("Data de Inicial da Suspensão maior que o cancelamento do Beneficiário !", "I")
        CheckVigenciaBenef = False
      End If
    End If
  End If
  Set SQL = Nothing
End Function

Public Function CHECARCOMPSUSPENSAO()

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  CHECARCOMPSUSPENSAO = True

  Condicao = " AND COMPETENCIAFINAL >=  COMPETENCIAINICIAL "

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_SUSPENSAO", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "BENEFICIARIO", Condicao)

  If Linha = "" Then
    CHECARCOMPSUSPENSAO = False
  Else
    CHECARCOMPSUSPENSAO = True
    bsShowMessage(Linha, "I")
  End If

  Set Interface = Nothing

End Function

