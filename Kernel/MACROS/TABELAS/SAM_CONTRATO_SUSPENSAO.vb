'HASH: 4939CBB05E51172C94BF590D3C81B42D
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
  q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  q1.Active = True

  vsValor = "3"

  If q1.FieldByName("CODIGO").AsInteger = 130 Then
    vsValor = "-1"
  End If

  Set q1 = Nothing

  If WebMode Then
	MOTIVO.WebLocalWhere = "TABTIPO <> "+vsValor+""     ' SMS 95929 - Paulo Melo - 17/04/2008 - Campos TAB devem ser comparados a INTEIROS para não dar problema em DB2
  ElseIf VisibleMode Then
  	MOTIVO.LocalWhere = "TABTIPO <> "+vsValor+""        ' SMS 95929 - Paulo Melo - 17/04/2008 - Campos TAB devem ser comparados a INTEIROS para não dar problema em DB2
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
  q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  q1.Active = True

  vsValor = "3"

  If q1.FieldByName("CODIGO").AsInteger = 130 Then
    vsValor = "-1"
  End If

  Set q1 = Nothing

  If WebMode Then
	MOTIVO.WebLocalWhere = "TABTIPO <> '"+vsValor+"'"
  ElseIf VisibleMode Then
  	MOTIVO.LocalWhere = "TABTIPO <> '"+vsValor+"'"
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Daniela -Início SMS 13047
  Dim q1 As Object
  Set q1 = NewQuery

  If (Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And(CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull)Then
    CanContinue = False
    bsShowMessage("A competência final não pode ser preenchida quando a competência inicial for nula.", "E")
    Exit Sub
  End If

  If(CurrentQuery.FieldByName("FATURARMESESSUSPENSOS").AsString = "S")And(CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)Then
    CanContinue = False
    bsShowMessage("Não é possível faturar meses suspensos sem que a competência final esteja preenchida.", "E")
    Exit Sub
  End If

  q1.Active = False
  q1.Clear
  q1.Add("SELECT TABTIPO, SUSPENDEFATURAMENTO FROM SAM_MOTIVOSUSPENSAO WHERE HANDLE = :HANDLE")
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


  q1.Active = False
  q1.Clear
  q1.Add("SELECT DATACANCELAMENTO FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
  q1.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  q1.Active = True

 If (q1.FieldByName("DATACANCELAMENTO").AsDateTime >= CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) And _
   (Not q1.FieldByName("DATACANCELAMENTO").IsNull)Then
  bsShowMessage("Não é possível suspender contrato, pois este já foi cancelado!", "E")
  CanContinue = False
 End If
 'FIm Daniela -SMS 13047

If Not(CurrentQuery.FieldByName("DATAFINAL").IsNull)And _
       (CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)Then
  CanContinue = False
  bsShowMessage("A data final não pode ser inferior à data inicial", "E")
  Exit Sub
End If


If Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull Then
  If Not(CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)And _
         (CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime <CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime)Then
    CanContinue = False
    bsShowMessage("A competência final não pode ser inferior à competência inicial", "E")
    Exit Sub
  End If
End If

Dim SQL As Object
Dim SQL1 As Object
Set SQL = NewQuery
Set SQL1 = NewQuery

SQL.Clear
SQL.Add("SELECT DATAADESAO")
SQL.Add("FROM SAM_CONTRATO")
SQL.Add("WHERE HANDLE = :HCONTRATO")
SQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
SQL.Active = True
If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
  CanContinue = False
  Set SQL = Nothing
  Set SQL1 = Nothing
  bsShowMessage("A data inicial não pode ser inferior à adesão do contrato", "E")
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
  bsShowMessage("A competência inicial não pode ser inferior à adesão do contrato", "E")
  Exit Sub
End If
End If

SQL.Clear
SQL.Add("SELECT * FROM SAM_MOTIVOSUSPENSAO WHERE HANDLE = :HMOTIVO")
SQL.ParamByName("HMOTIVO").AsInteger = CurrentQuery.FieldByName("MOTIVO").AsInteger
SQL.Active = True

SQL1.Add("SELECT * FROM SAM_CONTRATO WHERE HANDLE = :Hcontrato")
SQL1.ParamByName("Hcontrato").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
SQL1.Active = True


If(SQL.FieldByName("TABTIPO").AsInteger = 1 And SQL.FieldByName("NAOFATURARGUIAS").AsString = "S" _
   And SQL.FieldByName("NAOFATURARGUIAS").AsString = "S" And SQL1.FieldByName("LOCALFATURAMENTO").AsString = "F")Then
	bsShowMessage("Faturamento na Família, Opção Não Faturar Módulos e Não Faturar Guias marcadas!", "E")
	CanContinue = False

End If

SQL.Active = False
SQL1.Active = False

If CurrentQuery.State = 3 Then
  If CHECARSUSPENSAO Then
    CanContinue = False
    RefreshNodesWithTable("SAM_CONTRATO_SUSPENSAO")
    Exit Sub
  End If
End If

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


  If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
    vMesComp2 = DatePart("m", CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)
    vAnoComp2 = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime)
  End If

  vMesFechamento = DatePart("m", SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
  vAnoFechamento = DatePart("yyyy", SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

  If Not CurrentQuery.State = 3 Then
    If(vAnoComp1 <vAnoFechamento)Or _
       (vAnoComp1 = vAnoFechamento And vMesComp1 <vMesFechamento)Then
    CanContinue = False
    bsShowMessage("A competência inicial não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
  End If
End If
End If

If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
  If(vAnoComp2 <vAnoFechamento)Or _
     (vAnoComp2 = vAnoFechamento And vMesComp2 <vMesFechamento)Then
  CanContinue = False
  bsShowMessage("A competência final não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
End If
End If

Set SQLFechamento = Nothing


'sms 7934

Dim qTipoContrato
Dim qPessoa
Dim qSuspende
Dim qDependentes
Dim qContratos
Dim contratoSuspensos As Integer
Dim Contratos As String

Set qPessoa = NewQuery
Set qTipoContrato = NewQuery
Set qContratos = NewQuery
Set qSuspende = NewQuery
Set qDependentes = NewQuery


If CurrentQuery.State = 3 Then

  ContratosSuspensos = 0

  qTipoContrato.Add("SELECT TABTIPOCONTRATO,PESSOA FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
  qTipoContrato.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").Value
  qTipoContrato.Active = True


  qContratos.Add("SELECT HANDLE,CONTRATO,DATACANCELAMENTO FROM SAM_CONTRATO WHERE PESSOA = :PESSOA")


  qSuspende.Add("INSERT INTO SAM_CONTRATO_SUSPENSAO (HANDLE,CONTRATO,MOTIVO,OBSERVACAO,DATAINICIAL,DATAFINAL,COMPETENCIAINICIAL,COMPETENCIAFINAL)")
  qSuspende.Add("VALUES")
  qSuspende.Add("(:HANDLE,:CONTRATO,:MOTIVO,:OBSERVACAO,:DATAINICIAL,:DATAFINAL,:COMPETENCIAINICIAL,:COMPETENCIAFINAL)")


  If qTipoContrato.FieldByName("TABTIPOCONTRATO").AsInteger = 1 Then 'Contrato Empresarial

    qPessoa.Add("SELECT EHGRUPOFATURAMENTO FROM SFN_PESSOA WHERE HANDLE = :PESSOA")
    qPessoa.ParamByName("PESSOA").Value = qTipoContrato.FieldByName("PESSOA").AsInteger
    qPessoa.Active = True

    If qPessoa.FieldByName("EHGRUPOFATURAMENTO").AsString = "S" Then

      qDependentes.Add("SELECT HANDLE FROM SFN_PESSOA WHERE GRUPOFATURAMENTO = :PESSOA")
      qDependentes.ParamByName("PESSOA").Value = qTipoContrato.FieldByName("PESSOA").AsInteger
      qDependentes.Active = True

      If bsShowMessage("Cancelar contratos dependentes da pessoa responsável por este contrato?", "Q") = vbYes Then

        qSuspende.ParamByName("DATAFINAL").DataType = ftDateTime
        qSuspende.ParamByName("COMPETENCIAFINAL").DataType = ftDateTime

        While Not qDependentes.EOF

          qContratos.Active = False
          qContratos.ParamByName("PESSOA").Value = qDependentes.FieldByName("HANDLE").Value
          qContratos.Active = True

          'Contratos cancelados não devem ser suspensos
          If qContratos.FieldByName("DATACANCELAMENTO").IsNull Then

            While Not qContratos.EOF
              qSuspende.ParamByName("HANDLE").Value = NewHandle("SAM_CONTRATO_SUSPENSAO")
              qSuspende.ParamByName("CONTRATO").Value = qContratos.FieldByName("HANDLE").AsInteger
              qSuspende.ParamByName("MOTIVO").Value = CurrentQuery.FieldByName("MOTIVO").AsInteger
              qSuspende.ParamByName("OBSERVACAO").Value = CurrentQuery.FieldByName("OBSERVACAO").AsString
              qSuspende.ParamByName("DATAINICIAL").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
              If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
                qSuspende.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
              Else
                qSuspende.ParamByName("DATAFINAL").Value = Null
              End If

              If Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull Then
                qSuspende.ParamByName("COMPETENCIAINICIAL").Value = CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime
              Else
                qSuspende.ParamByName("COMPETENCIAINICIAL").Value = Null
              End If

              If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
                qSuspende.ParamByName("COMPETENCIAFINAL").Value = CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime
              Else
                qSuspende.ParamByName("COMPETENCIAFINAL").Value = Null
              End If
              qSuspende.ExecSQL
              qContratos.Next
              ContratosSuspensos = ContratosSuspensos + 1
              Contratos = "Contrato: " + qContratos.FieldByName("CONTRATO").AsString + Chr(13) + Contratos
            Wend
          End If
          qDependentes.Next
        Wend
      End If
    End If
    bsShowMessage("Foram suspensos " + ContratosSuspensos + " contratos" + Chr(13) + Contratos, "I")
  End If
End If

If CurrentQuery.State = 2 Then 'modo Edição

  If(Not CurrentQuery.FieldByName("DATAFINAL").IsNull)Or(Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull)Then
  If bsShowMessage("Finalizar suspensão dos contratos dependentes da pessoa responsável?", "Q") = vbYes Then
    Dim qGeral
    Set qGeral = NewQuery

    qGeral.Add("SELECT * FROM SAM_CONTRATO_SUSPENSAO WHERE (DATAINICIAL = :DATAINICIAL")

    If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
      qGeral.Add("AND COMPETENCIAINICIAL = :COMPETENCIAINICIAL ")
    End If

    qGeral.Add("AND MOTIVO = :MOTIVO and CONTRATO = :CONTRATO)")

    qSuspende.Add("UPDATE SAM_CONTRATO_SUSPENSAO SET DATAFINAL = :DATAFINAL, COMPETENCIAFINAL = :COMPETENCIAFINAL WHERE CONTRATO = :CONTRATO")

    qTipoContrato.Add("SELECT TABTIPOCONTRATO,PESSOA FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
    qTipoContrato.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").Value
    qTipoContrato.Active = True

    qContratos.Add("SELECT HANDLE,CONTRATO,DATACANCELAMENTO FROM SAM_CONTRATO WHERE PESSOA = :PESSOA")

    If qTipoContrato.FieldByName("TABTIPOCONTRATO").AsInteger = 1 Then 'Contrato Empresarial

      qPessoa.Add("SELECT EHGRUPOFATURAMENTO FROM SFN_PESSOA WHERE HANDLE = :PESSOA")
      qPessoa.ParamByName("PESSOA").Value = qTipoContrato.FieldByName("PESSOA").AsInteger
      qPessoa.Active = True

      If qPessoa.FieldByName("EHGRUPOFATURAMENTO").AsString = "S" Then

        qDependentes.Add("SELECT HANDLE FROM SFN_PESSOA WHERE GRUPOFATURAMENTO = :PESSOA")
        qDependentes.ParamByName("PESSOA").Value = qTipoContrato.FieldByName("PESSOA").AsInteger
        qDependentes.Active = True

        While Not qDependentes.EOF
          qContratos.Active = False
          qContratos.ParamByName("PESSOA").Value = qDependentes.FieldByName("HANDLE").Value
          qContratos.Active = True

          If qContratos.FieldByName("DATACANCELAMENTO").IsNull Then
            While Not qContratos.EOF
              qGeral.Active = False
              qGeral.ParamByName("DATAINICIAL").Value = CurrentQuery.FieldByName("DATAINICIAL").Value

              If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
                qGeral.ParamByName("COMPETENCIAINICIAL").Value = CurrentQuery.FieldByName("COMPETENCIAINICIAL").Value
              End If

              qGeral.ParamByName("MOTIVO").Value = CurrentQuery.FieldByName("MOTIVO").Value
              qGeral.ParamByName("CONTRATO").Value = qContratos.FieldByName("HANDLE").Value
              qGeral.Active = True

              If Not qGeral.EOF Then
                qSuspende.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").Value
                qSuspende.ParamByName("COMPETENCIAFINAL").Value = CurrentQuery.FieldByName("COMPETENCIAFINAL").Value
                qSuspende.ParamByName("CONTRATO").Value = qContratos.FieldByName("HANDLE").Value
                qSuspende.ExecSQL
              End If
              qContratos.Next
              ContratosSuspensos = ContratosSuspensos + 1
              Contratos = "Contrato: " + qContratos.FieldByName("CONTRATO").AsString + Chr(13) + Contratos
            Wend
          End If
          qDependentes.Next
        Wend
      End If

    End If
    Set qGeral = Nothing
  End If
End If
bsShowMessage("Foram finalizadas a suspensão de " + ContratosSuspensos + " contratos" + Chr(13) + Contratos, "I")
End If


Set qPessoa = Nothing
Set qTipoContrato = Nothing
Set qContratos = Nothing
Set qSuspende = Nothing
Set qDependentes = Nothing

End Sub

Public Function CHECARSUSPENSAO()
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  CHECARSUSPENSAO = True
  '---------------------------------
  'Ander sms 15717 13/05/03

  Condicao = "" ' " AND DATAFINAL >=  DATAINICIAL "
  '---------------------------------
  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_SUSPENSAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

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

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_SUSPENSAO", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha = "" Then
    CHECARCOMPSUSPENSAO = False
  Else
    CHECARCOMPSUSPENSAO = True
    bsShowMessage(Linha, "E")
  End If

  Set Interface = Nothing

End Function

