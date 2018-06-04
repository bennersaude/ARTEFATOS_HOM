'HASH: 81FCE3513363B160A1858079EF337FBD
'#Uses "*bsShowMessage"



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

  If Not CurrentQuery.FieldByName("ROTINASUSPENSAO").IsNull And CurrentQuery.State <> 3 Then
		OBSERVACAO.ReadOnly = True
  Else
		OBSERVACAO.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
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

  If WebMode Then
  	MOTIVO.WebLocalWhere = "TABTIPO <> "+vsValor+""  ' SMS 98104 - Paulo Melo - Campos TAB devem ser comparados a INTEIRO para não dar pau em DB2.
  ElseIf VisibleMode Then
  	MOTIVO.LocalWhere = "TABTIPO <> "+vsValor+""    ' SMS 98104 - Paulo Melo - Campos TAB devem ser comparados a INTEIRO para não dar pau em DB2.
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
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

  If WebMode Then
  	MOTIVO.WebLocalWhere = "TABTIPO <> "+vsValor+""    ' SMS 98104 - Paulo Melo - Campos TAB devem ser comparados a INTEIRO para não dar pau em DB2.
  ElseIf VisibleMode Then
  	MOTIVO.LocalWhere = "TABTIPO <> "+vsValor+""       ' SMS 98104 - Paulo Melo - Campos TAB devem ser comparados a INTEIRO para não dar pau em DB2.
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

If(CurrentQuery.FieldByName("FATURARMESESSUSPENSOS").AsString = "S")And(CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)Then
	CanContinue = False
	bsShowMessage("Não é possível faturar meses suspensos sem que a competência final esteja preenchida.", "E")
Exit Sub
End If

Dim q1 As Object
Set q1 = NewQuery



q1.Active = False
q1.Clear
q1.Add("SELECT SUSPENDEFATURAMENTO,TABTIPO FROM SAM_MOTIVOSUSPENSAO WHERE HANDLE = :HANDLE")
q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MOTIVO").AsInteger
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
SQL.Add("FROM SAM_FAMILIA")
SQL.Add("WHERE HANDLE = :HFAMILIA")
SQL.ParamByName("HFAMILIA").Value = CurrentQuery.FieldByName("FAMILIA").AsInteger
SQL.Active = True
If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
  CanContinue = False
  Set SQL = Nothing
  Set SQL1 = Nothing
  bsShowMessage("A data inicial não pode ser inferior à adesão da família", "E")
  Exit Sub
End If

Dim vMesComp As Integer
Dim vAnoComp As Integer
Dim vMesAdesao As Integer
Dim vAnoAdesao As Integer

If Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull Then
  vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)
  vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)
  vMesAdesao = DatePart("m", SQL.FieldByName("DATAADESAO").AsDateTime)
  vAnoAdesao = DatePart("yyyy", SQL.FieldByName("DATAADESAO").AsDateTime)

  If(vAnoComp <vAnoAdesao)Or _
     (vAnoComp = vAnoAdesao And vMesComp <vMesAdesao)Then
  CanContinue = False
  Set SQL = Nothing
  Set SQL1 = Nothing
  bsShowMessage("A competência inicial não pode ser inferior à adesão da família", "E")
  Exit Sub
End If
End If

SQL.Clear
SQL.Add("SELECT * FROM SAM_MOTIVOSUSPENSAO WHERE HANDLE = :HMOTIVO")
SQL.ParamByName("HMOTIVO").AsInteger = CurrentQuery.FieldByName("MOTIVO").AsInteger
SQL.Active = True

SQL1.Add("SELECT A.* FROM SAM_CONTRATO A, SAM_FAMILIA B WHERE A.HANDLE = B.CONTRATO AND B.HANDLE = :HFAMILIA")
SQL1.ParamByName("HFAMILIA").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
SQL1.Active = True


If(SQL.FieldByName("TABTIPO").AsInteger = 1 And SQL.FieldByName("NAOFATURARGUIAS").AsString = "S" _
   And SQL.FieldByName("NAOFATURARGUIAS").AsString = "S" And SQL1.FieldByName("LOCALFATURAMENTO").AsString = "C")Then
bsShowMessage("Faturamento no Contrato, Opção Não Faturar Módulos e Não Faturar Guias marcadas!", "E")
CanContinue = False

End If

SQL.Active = False
SQL1.Active = False

If CurrentQuery.State = 3 Then
  If CHECARSUSPENSAO Then
    CanContinue = False
    RefreshNodesWithTable("SAM_FAMILIA_SUSPENSAO")
    Exit Sub
  End If
End If

'If CHECARCOMPSUSPENSAO Then
'CanContinue =False
'RefreshNodesWithTable("SAM_FAMILIA_SUSPENSAO")
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

vMesFechamento = DatePart("m", SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
vAnoFechamento = DatePart("yyyy", SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

If Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull Then
  If CurrentQuery.State = 3 Then
    vMesComp1 = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)
    vAnoComp1 = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)

    If(vAnoComp1 <vAnoAdesao)Or _
       (vAnoComp1 = vAnoAdesao And vMesComp1 <vMesFechamento)Then
    CanContinue = False
    bsShowMessage("A competência inicial não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
  End If
End If
End If


vMesComp2 = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)
vAnoComp2 = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)

If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
  If(vAnoComp2 <vAnoAdesao)Or _
     (vAnoComp2 = vAnoAdesao And vMesComp2 <vMesFechamento)Then
  CanContinue = False
  bsShowMessage("A competência final não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
End If
End If

Set SQLFechamento = Nothing
Set q1 = Nothing
'*****************************************************************************************************************

'*****************************************************************************************************************
End Sub

Public Function CHECARSUSPENSAO()
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim SQL As Object

  CHECARSUSPENSAO = True
  '---------------------------------------------
  'Ander sms 15717 13/05/03

  Condicao = "" 'Condicao =" AND DATAFINAL >=  DATAINICIAL "
  '---------------------------------------------
  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_FAMILIA_SUSPENSAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "FAMILIA", Condicao)

  If Linha = "" Then
    CHECARSUSPENSAO = False
  Else
    CHECARSUSPENSAO = True
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing

End Function

Public Function CHECARCOMPSUSPENSAO()

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  CHECARCOMPSUSPENSAO = True

  Condicao = " AND COMPETENCIAFINAL >=  COMPETENCIAINICIAL "

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_FAMILIA_SUSPENSAO", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "FAMILIA", Condicao)

  If Linha = "" Then
    CHECARCOMPSUSPENSAO = False
  Else
    CHECARCOMPSUSPENSAO = True
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing

End Function

