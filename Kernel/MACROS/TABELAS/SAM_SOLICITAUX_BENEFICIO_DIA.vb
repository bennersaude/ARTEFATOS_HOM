'HASH: 6D9BB3EA546360FDB4202E1EABB8E3F8

'SAM_SOLICITAUX_BENEFICIO_DIA
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterDelete()
  RefreshNodesWithTable("SAM_SOLICITAUX")
End Sub

Public Sub TABLE_AfterInsert()
  Dim SQLDOC As Object
  Dim SQLDIA
  Set SQLDOC = NewQuery
  Set SQLDIA = NewQuery

  SQLDIA.Active = False
  SQLDIA.Clear
  SQLDIA.Add("SELECT SOLICITAUX, VALORPRESTCONTAS FROM SAM_SOLICITAUX_BENEFICIO WHERE HANDLE = :HSOLICITAUXBEN")
  SQLDIA.ParamByName("HSOLICITAUXBEN").AsInteger = RecordHandleOfTable("SAM_SOLICITAUX_BENEFICIO")
  SQLDIA.Active = True

  SQLDOC.Active = False
  SQLDOC.Clear
  SQLDOC.Add("SELECT SITUACAO FROM SAM_SOLICITAUX WHERE HANDLE = :HSOLICITAUX")
  SQLDOC.ParamByName("HSOLICITAUX").AsInteger = SQLDIA.FieldByName("SOLICITAUX").AsInteger
  SQLDOC.Active = True

  If SQLDOC.FieldByName("SITUACAO").AsString <> "A" Then
    QTDDIARIASSOLIC.ReadOnly = True
    VALORHOSPEDAGEMSOLIC.ReadOnly = True
    VALORREFEICAOSOLIC.ReadOnly = True

    If SQLDOC.FieldByName("SITUACAO").AsString = "L" And _
                          SQLDIA.FieldByName("VALORPRESTCONTAS").IsNull Then
      TIPODIARIA.ReadOnly = False
      QTDDIARIASPRESTCONTAS.ReadOnly = False
      VALORHOSPEDAGEMPRESTCONTAS.ReadOnly = False
      VALORREFEICAOPRESTCONTAS.ReadOnly = False
    Else
      TIPODIARIA.ReadOnly = True
      QTDDIARIASPRESTCONTAS.ReadOnly = True
      VALORHOSPEDAGEMPRESTCONTAS.ReadOnly = True
      VALORREFEICAOPRESTCONTAS.ReadOnly = True
    End If
  Else
    QTDDIARIASSOLIC.ReadOnly = False
    QTDDIARIASPRESTCONTAS.ReadOnly = False
    TIPODIARIA.ReadOnly = False
    VALORHOSPEDAGEMSOLIC.ReadOnly = False
    VALORHOSPEDAGEMPRESTCONTAS.ReadOnly = False
    VALORREFEICAOSOLIC.ReadOnly = False
    VALORREFEICAOPRESTCONTAS.ReadOnly = False
  End If
  SQLDIA.Active = False
  SQLDOC.Active = False
  Set SQLDIA = Nothing
  Set SQLDOC = Nothing
End Sub

Public Sub TABLE_AfterScroll()

  Dim SQLVER As Object
  Dim SQLDOC As Object
  Dim SQLDIA As Object
  Set SQLVER = NewQuery
  Set SQLDOC = NewQuery
  Set SQLDIA = NewQuery

  SQLDIA.Active = False
  SQLDIA.Clear
  SQLDIA.Add("SELECT SOLICITAUX, VALORPRESTCONTAS FROM SAM_SOLICITAUX_BENEFICIO WHERE HANDLE = :HSOLICITAUXBEN")
  SQLDIA.ParamByName("HSOLICITAUXBEN").AsInteger = RecordHandleOfTable("SAM_SOLICITAUX_BENEFICIO")
  SQLDIA.Active = True

  SQLDOC.Active = False
  SQLDOC.Clear
  SQLDOC.Add("SELECT SITUACAO FROM SAM_SOLICITAUX WHERE HANDLE = :HSOLICITAUX")
  SQLDOC.ParamByName("HSOLICITAUX").AsInteger = SQLDIA.FieldByName("SOLICITAUX").AsInteger
  SQLDOC.Active = True

  If SQLDOC.FieldByName("SITUACAO").AsString <> "A" Then
    TIPODIARIA.ReadOnly = True
    QTDDIARIASSOLIC.ReadOnly = True
    VALORHOSPEDAGEMSOLIC.ReadOnly = True
    VALORREFEICAOSOLIC.ReadOnly = True

    If SQLDOC.FieldByName("SITUACAO").AsString = "L" And _
                          SQLDIA.FieldByName("VALORPRESTCONTAS").IsNull Then
      QTDDIARIASPRESTCONTAS.ReadOnly = False
      VALORHOSPEDAGEMPRESTCONTAS.ReadOnly = False
      VALORREFEICAOPRESTCONTAS.ReadOnly = False
    Else
      QTDDIARIASPRESTCONTAS.ReadOnly = True
      VALORHOSPEDAGEMPRESTCONTAS.ReadOnly = True
      VALORREFEICAOPRESTCONTAS.ReadOnly = True
    End If
  Else
    QTDDIARIASSOLIC.ReadOnly = False
    QTDDIARIASPRESTCONTAS.ReadOnly = False
    TIPODIARIA.ReadOnly = False
    VALORHOSPEDAGEMSOLIC.ReadOnly = False
    VALORHOSPEDAGEMPRESTCONTAS.ReadOnly = False
    VALORREFEICAOSOLIC.ReadOnly = False
    VALORREFEICAOPRESTCONTAS.ReadOnly = False
  End If

  SQLVER.Active = False
  SQLVER.Clear
  SQLVER.Add(" SELECT COUNT(A.HANDLE) AS TOTAL")
  SQLVER.Add(" FROM SAM_SOLICITAUX_BENEFICIO A, ")
  SQLVER.Add("      SAM_SOLICITAUX_BENEFICIO_DIA B")
  SQLVER.Add(" WHERE (VALORPRESTCONTAS IS NOT NULL Or VALORPRESTCONTAS <> 0)")
  SQLVER.Add("   AND (A.HANDLE = :HSOLICITAUX)")
  SQLVER.Add("   AND B.SOLICITAUXBENEFICIO = A.HANDLE")
  SQLVER.ParamByName("HSOLICITAUX").AsInteger = CurrentQuery.FieldByName("SOLICITAUXBENEFICIO").AsInteger
  SQLVER.Active = True

  If SQLVER.FieldByName("TOTAL").AsInteger <> 0 Then
    BsShowMessage("Não é possível alterar, prestação de contas efetuada!", "I")
    RefreshNodesWithTable("SAM_SOLICITAUX_BENEFICIO_DIA")
    Exit Sub
  End If

  SQLVER.Active = False
  SQLDOC.Active = False
  SQLDIA.Active = False
  Set SQLVER = Nothing
  Set SQLDIA = Nothing
  Set SQLDOC = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  TIPODIARIA.ReadOnly = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLDIA As Object
  Dim SQLSUM As Object
  Dim SQLTDIA As Object
  Set SQLDIA = NewQuery
  Set SQLSUM = NewQuery
  Set SQLTDIA = NewQuery


  SQLSUM.Active = False
  SQLSUM.Clear
  SQLSUM.Add(" SELECT SUM(QTDDIARIASSOLIC) AS TOTALDIARIAS,  ")
  SQLSUM.Add(" SUM(VALORREFEICAOSOLIC) AS TOTALREFEICAO,     ")
  SQLSUM.Add(" SUM(VALORHOSPEDAGEMSOLIC) AS TOTALHOSPEDAGEM  ")
  SQLSUM.Add(" FROM SAM_SOLICITAUX_BENEFICIO_DIA             ")
  SQLSUM.Add(" WHERE HANDLE <> :HBENEFDIA                    ")
  SQLSUM.Add("   AND SOLICITAUXBENEFICIO = :HBENEFICIO       ")
  SQLSUM.Add("   AND TIPODIARIA = :HTIPODIARIA               ")
  SQLSUM.ParamByName("HBENEFDIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLSUM.ParamByName("HBENEFICIO").Value = RecordHandleOfTable("SAM_SOLICITAUX_BENEFICIO")
  SQLSUM.ParamByName("HTIPODIARIA").Value = CurrentQuery.FieldByName("TIPODIARIA").AsInteger
  SQLSUM.Active = True

  SQLTDIA.Active = False
  SQLTDIA.Clear
  SQLTDIA.Add(" SELECT LIMITEDIARIAS, VALORHOSPEDAGEM, VALORREFEICAO FROM SAM_TIPODIARIA")
  SQLTDIA.Add(" WHERE HANDLE = :TIPODIARIA")
  SQLTDIA.ParamByName("TIPODIARIA").AsInteger = CurrentQuery.FieldByName("TIPODIARIA").AsInteger
  SQLTDIA.Active = True

  If (CurrentQuery.FieldByName("QTDDIARIASSOLIC").AsInteger + _
      SQLSUM.FieldByName("TOTALDIARIAS").AsInteger) > SQLTDIA.FieldByName("LIMITEDIARIAS").AsInteger Then

	BsShowMessage("A quantidade de diárias ultrapassou o limite", "I")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("VALORHOSPEDAGEMSOLIC").AsCurrency + _
      SQLSUM.FieldByName("TOTALHOSPEDAGEM").AsCurrency) > (CurrentQuery.FieldByName("QTDDIARIASSOLIC").AsInteger * SQLTDIA.FieldByName("VALORHOSPEDAGEM").AsCurrency) Then
    BsShowMessage("O valor total das hospedagens ultrapassou o limite", "I")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("VALORREFEICAOSOLIC").AsCurrency + _
      SQLSUM.FieldByName("TOTALREFEICAO").AsCurrency) > (CurrentQuery.FieldByName("QTDDIARIASSOLIC").AsInteger * SQLTDIA.FieldByName("VALORREFEICAO").AsCurrency) Then
    BsShowMessage("O valor total das refeições ultrapassou o limite", "I")
    CanContinue = False
    Exit Sub
  End If

End Sub

