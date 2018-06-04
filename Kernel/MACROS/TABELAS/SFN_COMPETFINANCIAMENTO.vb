'HASH: C8B5B92FB74C3A1E72CB8ADCDD7FB566
Option Explicit

Public Sub TABLE_AfterInsert()
  Dim sql As Object
  Set sql = NewQuery

  Dim sql2 As Object
  Set sql2 = NewQuery

  sql.Active = False
  sql.Add("SELECT MAX(COMPETENCIA) REC FROM SFN_COMPETFINANCIAMENTO")
  sql.Active = True

  sql2.Active = False
  sql2.Add("SELECT COMPETENCIA, ")
  sql2.Add("       DIAFINAL ")
  sql2.Add("  FROM SFN_COMPETFINANCIAMENTO ")
  sql2.Add(" WHERE COMPETENCIA = :COMPET ")
  sql2.ParamByName("COMPET").Value = sql.FieldByName("REC").AsDateTime
  sql2.Active = True

  If Not sql2.EOF Then

    Dim viMes As Integer
    Dim viAno As Integer
    Dim viDia As Integer
    Dim viDiaAux As Integer

    viMes = Month(sql2.FieldByName("COMPETENCIA").AsDateTime)
    viAno = Year(sql2.FieldByName("COMPETENCIA").AsDateTime)
    '    viDia = CurrentQuery.FieldByName("DIAFINAL").AsInteger
    If (viMes = 1) Or (viMes = 3) Or (viMes = 5) Or (viMes = 7) Or (viMes = 8) Or (viMes = 10) Or (viMes = 12) Then
      viDia = 31
    Else
      If viMes <> 2 Then
        viDia = 30
      Else
        If (Int(viAno) Mod 4) = 0 Then
          viDia = 29
        Else
          viDia = 28
        End If
      End If
    End If

    CurrentQuery.FieldByName("COMPETENCIA").AsDateTime = DateAdd("m", 1, sql2.FieldByName("COMPETENCIA").AsDateTime)

    If sql2.FieldByName("DIAFINAL").AsInteger + 1 > viDia Then
      CurrentQuery.FieldByName("DIAINICIAL").AsInteger = 1
    Else
      CurrentQuery.FieldByName("DIAINICIAL").AsInteger = sql2.FieldByName("DIAFINAL").AsInteger + 1
    End If
    COMPETENCIA.ReadOnly = True
    DIAINICIAL.ReadOnly = True
  Else
    COMPETENCIA.ReadOnly = False
    DIAINICIAL.ReadOnly = False
  End If

  Set sql = Nothing
  Set sql2 = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("DIAINICIAL").AsInteger > 1 Then
    CurrentQuery.FieldByName("DATAINICIAL").AsDateTime = DateAdd("m", -1, CurrentQuery.FieldByName("COMPETENCIA").AsDateTime) + (CurrentQuery.FieldByName("DIAINICIAL").AsInteger -1)
  Else
    CurrentQuery.FieldByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime' + (CurrentQuery.FieldByName("DIAINICIAL").AsInteger)
  End If

  Dim viMes As Integer
  Dim viAno As Integer
  Dim viDia As Integer
  Dim viDiaAux As Integer

  viMes = Month(CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
  viAno = Year(CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
  viDia = CurrentQuery.FieldByName("DIAFINAL").AsInteger
  If (viMes = 1) Or (viMes = 3) Or (viMes = 5) Or (viMes = 7) Or (viMes = 8) Or (viMes = 10) Or (viMes = 12) Then
    viDia = 31
  Else
    If viMes <> 2 Then
      viDia = 30
    Else
      If (Int(viAno) Mod 4) = 0 Then
        viDia = 29
      Else
        viDia = 28
      End If
    End If
  End If

  If CurrentQuery.FieldByName("DIAFINAL").AsInteger > viDia Then
    viDiaAux = Day(CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
    CurrentQuery.FieldByName("DATAFINAL").AsDateTime = (CurrentQuery.FieldByName("COMPETENCIA").AsDateTime + viDia) - viDiaAux
    CurrentQuery.FieldByName("DIAFINAL").AsInteger = viDia
  Else
    CurrentQuery.FieldByName("DATAFINAL").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime + CurrentQuery.FieldByName("DIAFINAL").AsInteger -1
  End If



End Sub



