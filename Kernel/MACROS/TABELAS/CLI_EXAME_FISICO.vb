'HASH: 23FFA7B6D21DC63A76EAF87029724FB4
'#Uses "*bsShowMessage"

Option Explicit

Dim EXAME_FISICO As Boolean


Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("SUBJETIVO").Value = CLng(SessionVar("CLI_SUBJETIVO"))
  CurrentQuery.FieldByName("PROFISSIONAL").Value = CLng(SessionVar("CLI_RECURSO"))
  CurrentQuery.FieldByName("ESPECIALIDADE").Value = CLng(SessionVar("SAM_ESPECIALIDADE"))
  CurrentQuery.FieldByName("PACIENTE").Value = CLng(SessionVar("SAM_MATRICULA"))
  CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger = CurrentUser
  CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime = ServerDate

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False

  SQL.Clear
  SQL.Add("SELECT DATAABERTURA     ")
  SQL.Add("  FROM CLI_SUBJETIVO    ")
  SQL.Add(" WHERE HANDLE = :HANDLE ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("SUBJETIVO").AsInteger
  SQL.Active = True

  CurrentQuery.FieldByName("IDADE").AsInteger = CalculaIdadeFuncao(CurrentQuery.FieldByName("PACIENTE").AsInteger, SQL.FieldByName("DATAABERTURA").AsDateTime)

  SQL.Active = False
  Set SQL = Nothing

End Sub


Function CalculaIdadeFuncao(ByVal piMatricula As Long, ByVal pdDataAplicacao As Date) As Integer
  Dim viDias           As Integer
  Dim viMeses          As Integer
  Dim viAnos           As Integer
  Dim VdDataNascimento As Date

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT DATANASCIMENTO   ")
  SQL.Add("  FROM SAM_MATRICULA    ")
  SQL.Add(" WHERE HANDLE = :HANDLE ")
  SQL.ParamByName("HANDLE").AsInteger = piMatricula
  SQL.Active = True

  VdDataNascimento = SQL.FieldByName("DATANASCIMENTO").AsDateTime

  If (VdDataNascimento > ServerDate) Then
    CalculaIdadeFuncao = 0
  Else
    DiferencaDataFuncao pdDataAplicacao, VdDataNascimento, viDias, viMeses, viAnos
    CalculaIdadeFuncao = (viAnos * 12) + viMeses
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function


Public Sub DiferencaDataFuncao(ByVal Data1, Data2 As Date, Dias, Meses, Anos As Integer)
  Dim DtSwap As Date
  Dim Day1, Day2, Month1, Month2, Year1, Year2 As Integer

  If Data1 >Data2 Then
    DtSwap = Data1
    Data1 = Data2
    Data2 = DtSwap
  End If

  Year1 = Val(Format(Data1, "yyyy"))
  Month1 = Val(Format(Data1, "mm"))
  Day1 = Val(Format(Data1, "dd"))

  Year2 = Val(Format(Data2, "yyyy"))
  Month2 = Val(Format(Data2, "mm"))
  Day2 = Val(Format(Data2, "dd"))

  Anos = Year2 - Year1
  Meses = 0
  Dias = 0

  If Month2 <Month1 Then
    Meses = Meses + 12
    Anos = Anos -1
  End If

  Meses = Meses + (Month2 - Month1)

  If Day2 <Day1 Then
    Dias = Dias + DiasPorMesFuncao(Year1, Val(Month1))
    If Meses = 0 Then
      Anos = Anos - 1
      Meses = 11
    Else
      Meses = Meses -1
    End If
  End If
  Dias = Dias + (Day2 - Day1)
End Sub

Function DiasPorMesFuncao(ByVal Ano, Mes As Integer)As Integer
  Dim Meses31 As String
  Dim Meses30 As String

  Meses31 = "'1','3','5','7','8','10','12'"
  Meses30 = "'4','6','9','11'"

  If InStr(Meses31, "'" + Str(Mes) + "'") > 0 Then
    DiasPorMesFuncao = 31
  ElseIf InStr(Meses30, "'" + Str(Mes) + "'") > 0 Then
    DiasPorMesFuncao = 30
  Else
    If Ano Mod 4 = 0 Then
      DiasPorMesFuncao = 29
    Else
      DiasPorMesFuncao = 28
    End If
  End If

End Function
