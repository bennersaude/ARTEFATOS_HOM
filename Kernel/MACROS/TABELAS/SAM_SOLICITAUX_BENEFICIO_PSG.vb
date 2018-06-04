'HASH: 1C16AED9C85A8087B97283F74FD77CE7
'SAM_SOLICITAUX_BENEFICIO_PSG
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
  Dim SQLDOC As Object
  Dim SQLDIA
  Set SQLDOC = NewQuery
  Set SQLDIA = NewQuery

  SQLDIA.Active = False
  SQLDIA.Clear
  SQLDIA.Add("SELECT SOLICITAUX, VALORPRESTCONTAS FROM SAM_SOLICITAUX_BENEFICIO WHERE HANDLE = :HSOLICITAUXBEN")
  SQLDIA.ParamByName("HSOLICITAUXBEN").AsInteger = CurrentQuery.FieldByName("SOLICITAUXBENEFICIO").AsInteger
  SQLDIA.Active = True

  SQLDOC.Active = False
  SQLDOC.Clear
  SQLDOC.Add("SELECT SITUACAO FROM SAM_SOLICITAUX WHERE HANDLE = :HSOLICITAUX")
  SQLDOC.ParamByName("HSOLICITAUX").AsInteger = SQLDIA.FieldByName("SOLICITAUX").AsInteger
  SQLDOC.Active = True

  If SQLDOC.FieldByName("SITUACAO").AsString <> "A" Then
    QTDPASSAGENSSOLIC.ReadOnly = True
    VALORTOTALSOLIC.ReadOnly = True

    If SQLDOC.FieldByName("SITUACAO").AsString = "L" And _
                          SQLDIA.FieldByName("VALORPRESTCONTAS").IsNull Then
      TIPODESLOCAMENTO.ReadOnly = False
      QTDPASSAGENSPRESTCONTAS.ReadOnly = False
      VALORTOTALPRESTCONTAS.ReadOnly = False
    Else
      TIPODESLOCAMENTO.ReadOnly = True
      QTDPASSAGENSPRESTCONTAS.ReadOnly = True
      VALORTOTALPRESTCONTAS.ReadOnly = True
    End If
  Else
    QTDPASSAGENSPRESTCONTAS.ReadOnly = False
    QTDPASSAGENSSOLIC.ReadOnly = False
    TIPODESLOCAMENTO.ReadOnly = False
    VALORTOTALSOLIC.ReadOnly = False
    VALORTOTALPRESTCONTAS.ReadOnly = False
  End If
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQLDOC As Object
  Dim SQLDIA As Object
  Dim SQLVER As Object
  Set SQLDOC = NewQuery
  Set SQLDIA = NewQuery
  Set SQLVER = NewQuery

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
    QTDPASSAGENSSOLIC.ReadOnly = True
    VALORTOTALSOLIC.ReadOnly = True
    TIPODESLOCAMENTO.ReadOnly = True

    If SQLDOC.FieldByName("SITUACAO").AsString = "L" And _
                          SQLDIA.FieldByName("VALORPRESTCONTAS").IsNull Then
      TIPODESLOCAMENTO.ReadOnly = False
      QTDPASSAGENSPRESTCONTAS.ReadOnly = False
      VALORTOTALPRESTCONTAS.ReadOnly = False
    Else
      TIPODESLOCAMENTO.ReadOnly = True
      QTDPASSAGENSPRESTCONTAS.ReadOnly = True
      VALORTOTALPRESTCONTAS.ReadOnly = True
    End If
  Else
    QTDPASSAGENSPRESTCONTAS.ReadOnly = False
    QTDPASSAGENSSOLIC.ReadOnly = False
    TIPODESLOCAMENTO.ReadOnly = False
    VALORTOTALSOLIC.ReadOnly = False
    VALORTOTALPRESTCONTAS.ReadOnly = False
  End If


  SQLVER.Active = False
  SQLVER.Clear
  SQLVER.Add(" SELECT COUNT(A.HANDLE) AS TOTAL")
  SQLVER.Add(" FROM SAM_SOLICITAUX_BENEFICIO A, ")
  SQLVER.Add("      SAM_SOLICITAUX_BENEFICIO_DIA B")
  SQLVER.Add(" WHERE (VALORPRESTCONTAS IS NOT NULL Or VALORPRESTCONTAS <> 0)")
  SQLVER.Add("   AND A.HANDLE = :HSOLICITAUX")
  SQLVER.Add("   AND B.SOLICITAUXBENEFICIO = A.HANDLE")
  SQLVER.ParamByName("HSOLICITAUX").AsInteger = CurrentQuery.FieldByName("SOLICITAUXBENEFICIO").AsInteger
  SQLVER.Active = True

  If SQLVER.FieldByName("TOTAL").AsInteger <> 0 Then
  	BsShowMessage("Não é possível alterar, prestação de contas efetuada!", "I")
    RefreshNodesWithTable("SAM_SOLICITAUX_BENEFICIO_DIA")
    Exit Sub
  End If

  SQLDOC.Active = False
  SQLDIA.Active = False
  SQLVER.Active = False
  Set SQLDOC = Nothing
  Set SQLDIA = Nothing
  Set SQLVER = Nothing
End Sub

