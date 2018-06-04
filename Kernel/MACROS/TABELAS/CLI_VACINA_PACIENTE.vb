'HASH: 104737E780BE714E43425F13E7C23E2B
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("MATRICULA").AsInteger = CLng(SessionVar("SAM_MATRICULA"))
  CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger = CurrentUser
End Sub

Public Sub TABLE_AfterPost()
  If (CurrentQuery.FieldByName("DOSE").AsInteger < 6) Then
    Verificar_E_RegistrarRetorno(CurrentQuery.FieldByName("HANDLE").AsInteger)
  End If
End Sub

Function CalculaIdade(ByVal piMatricula As Long, ByVal pdDataAplicacao As Date) As Integer
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
    CalculaIdade = 0
  Else
    DiferencaData pdDataAplicacao, VdDataNascimento, viDias, viMeses, viAnos
    CalculaIdade = (viAnos * 12) + viMeses
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Sub DiferencaData(ByVal Data1, Data2 As Date, Dias, Meses, Anos As Integer)
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
    Dias = Dias + DiasPorMes(Year1, Val(Month1))
    If Meses = 0 Then
      Anos = Anos - 1
      Meses = 11
    Else
      Meses = Meses -1
    End If
  End If
  Dias = Dias + (Day2 - Day1)
End Sub

Function DiasPorMes(ByVal Ano, Mes As Integer)As Integer
  Dim Meses31 As String
  Dim Meses30 As String

  Meses31 = "'1','3','5','7','8','10','12'"
  Meses30 = "'4','6','9','11'"

  If InStr(Meses31, "'" + Str(Mes) + "'") > 0 Then
    DiasPorMes = 31
  ElseIf InStr(Meses30, "'" + Str(Mes) + "'") > 0 Then
    DiasPorMes = 30
  Else
    If Ano Mod 4 = 0 Then
      DiasPorMes = 29
    Else
      DiasPorMes = 28
    End If
  End If

End Function

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  'Se o registro que está sendo excluído tem retorno posterior, então o mesmo deve ser excluído também
  SQL.Active = False
  SQL.Clear
  SQL.Add("DELETE                        ")
  SQL.Add("  FROM CLI_VACINA_PACIENTE    ")
  SQL.Add(" WHERE VACINA = :VACINA       ")
  SQL.Add("   AND MATRICULA = :MATRICULA ")
  SQL.Add("   AND HANDLE > :HANDLE       ")
  SQL.Add("   AND DATAAPLICACAO IS NULL  ")
  SQL.ParamByName("MATRICULA").AsInteger = CurrentQuery.FieldByName("MATRICULA").AsInteger
  SQL.ParamByName("VACINA").AsInteger    = CurrentQuery.FieldByName("VACINA").AsInteger
  SQL.ParamByName("HANDLE").AsInteger    = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL

  'Verifica qual o último registro da VACINA que restou para calcular a nova data de retorno, se houver retorno
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT MAX(HANDLE) MHANDLE    ")
  SQL.Add("  FROM CLI_VACINA_PACIENTE    ")
  SQL.Add(" WHERE VACINA    = :VACINA    ")
  SQL.Add("   AND MATRICULA = :MATRICULA ")
  SQL.Add("   AND HANDLE   <> :HANDLE    ")
  SQL.ParamByName("MATRICULA").AsInteger = CurrentQuery.FieldByName("MATRICULA").AsInteger
  SQL.ParamByName("VACINA").AsInteger    = CurrentQuery.FieldByName("VACINA").AsInteger
  SQL.ParamByName("HANDLE").AsInteger    = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If SQL.FieldByName("MHANDLE").AsInteger > 0 Then
    Verificar_E_RegistrarRetorno(SQL.FieldByName("MHANDLE").AsInteger)
  End If

  SQL.Active = False
  Set SQL = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("MATRICULA").AsInteger > 0 Then
    Dim viIdadeMeses As Integer
    viIdadeMeses = CalculaIdade(CurrentQuery.FieldByName("MATRICULA").AsInteger, CurrentQuery.FieldByName("DATAAPLICACAO").AsDateTime)
    If viIdadeMeses > 24 Then
      CurrentQuery.FieldByName("IDADE").AsInteger = CInt(viIdadeMeses/12)
      CurrentQuery.FieldByName("UNIDADEIDADE").AsInteger = 2
    Else
      CurrentQuery.FieldByName("IDADE").AsInteger = viIdadeMeses
      CurrentQuery.FieldByName("UNIDADEIDADE").AsInteger = 1
    End If
  End If

  If CurrentQuery.State = 3 Then 'Se for inclusão, calcula a última dose
    CurrentQuery.FieldByName("DOSE").AsInteger = ObterUltimaDose + 1
  End If

End Sub

Function ObterUltimaDose As Integer
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT MAX(DOSE) ULTIMADOSE   ")
  SQL.Add("  FROM CLI_VACINA_PACIENTE    ")
  SQL.Add(" WHERE MATRICULA = :MATRICULA ")
  SQL.Add("   AND VACINA    = :VACINA    ")
  SQL.ParamByName("MATRICULA").AsInteger = CurrentQuery.FieldByName("MATRICULA").AsInteger
  SQL.ParamByName("VACINA").AsInteger    = CurrentQuery.FieldByName("VACINA").AsInteger
  SQL.Active = True
  ObterUltimaDose = SQL.FieldByName("ULTIMADOSE").AsInteger
  SQL.Active = False
  Set SQL = Nothing
End Function

Public Sub Verificar_E_RegistrarRetorno(ByVal piHandleAplicacao As Long)

  Dim vdDataRetorno As Date
  Dim vdDataAplicacao As Date
  Dim viDose As Integer

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT DATAAPLICACAO,      ")
  SQL.Add("       DOSE                ")
  SQL.Add("  FROM CLI_VACINA_PACIENTE ")
  SQL.Add(" WHERE HANDLE = :HANDLE    ")
  SQL.ParamByName("HANDLE").AsInteger = piHandleAplicacao
  SQL.Active = True

  vdDataAplicacao = SQL.FieldByName("DATAAPLICACAO").AsDateTime
  viDose = SQL.FieldByName("DOSE").AsInteger

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT *                ")
  SQL.Add("  FROM CLI_VACINA       ")
  SQL.Add(" WHERE HANDLE = :HANDLE ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("VACINA").AsInteger
  SQL.Active = True

  If (SQL.FieldByName("TABDOSEUNICA").AsInteger = 1) Then
    Exit Sub
  ElseIf (viDose = 1) And (SQL.FieldByName("DOSE2").AsInteger = 1) Then   'Primeira dose e existe a segunda
    vdDataRetorno = ObterDataRetorno(vdDataAplicacao, _
                                     SQL.FieldByName("TEMPORETORNODOSE2").AsInteger, _
                                     SQL.FieldByName("UNIDADERETORNODOSE2").AsInteger)
    InserirRegistroRetorno(piHandleAplicacao, _
                           CurrentQuery.FieldByName("MATRICULA").AsInteger, _
                           CurrentQuery.FieldByName("VACINA").AsInteger, _
                           2, _
                           vdDataRetorno)
  ElseIf (viDose = 2) And (SQL.FieldByName("DOSE3").AsInteger = 1) Then   'Segunda dose e existe a terceira
    vdDataRetorno = ObterDataRetorno(vdDataAplicacao, _
                                     SQL.FieldByName("TEMPORETORNODOSE3").AsInteger, _
                                     SQL.FieldByName("UNIDADERETORNODOSE3").AsInteger)
    InserirRegistroRetorno(piHandleAplicacao, _
                           CurrentQuery.FieldByName("MATRICULA").AsInteger, _
                           CurrentQuery.FieldByName("VACINA").AsInteger, _
                           3, _
                           vdDataRetorno)
  ElseIf (viDose = 3) And (SQL.FieldByName("REFORCO").AsInteger = 1) Then 'Terceira dose e existe a quarta
    vdDataRetorno = ObterDataRetorno(vdDataAplicacao, _
                                     SQL.FieldByName("TEMPORETORNOREFORCO").AsInteger, _
                                     SQL.FieldByName("UNIDADERETORNOREFORCO").AsInteger)
    InserirRegistroRetorno(piHandleAplicacao, _
                           CurrentQuery.FieldByName("MATRICULA").AsInteger, _
                           CurrentQuery.FieldByName("VACINA").AsInteger, _
                           4, _
                           vdDataRetorno)
  ElseIf (viDose = 4) And (SQL.FieldByName("REFORCO").AsInteger = 1) Then 'Reforço e existe retorno para novo reforço
    vdDataRetorno = ObterDataRetorno(vdDataAplicacao, _
                                     SQL.FieldByName("TEMPORETORNOREFORCO").AsInteger, _
                                     SQL.FieldByName("UNIDADERETORNOREFORCO").AsInteger)
    InserirRegistroRetorno(piHandleAplicacao, _
                           CurrentQuery.FieldByName("MATRICULA").AsInteger, _
                           CurrentQuery.FieldByName("VACINA").AsInteger, _
                           5, _
                           vdDataRetorno)
  ElseIf (viDose = 5) And (SQL.FieldByName("REFORCO").AsInteger = 1) Then 'Reforço e existe retorno para novo reforço
    vdDataRetorno = ObterDataRetorno(vdDataAplicacao, _
                                     SQL.FieldByName("TEMPORETORNOREFORCO").AsInteger, _
                                     SQL.FieldByName("UNIDADERETORNOREFORCO").AsInteger)
    InserirRegistroRetorno(piHandleAplicacao, _
                           CurrentQuery.FieldByName("MATRICULA").AsInteger, _
                           CurrentQuery.FieldByName("VACINA").AsInteger, _
                           6, _
                           vdDataRetorno)
  End If

  SQL.Active = False
  Set SQL = Nothing
End Sub

Public Function ObterDataRetorno(ByVal pdDataAplicacao As Date, ByVal piTempoRetorno As Integer, ByVal piUnidadeRetorno As Integer) As Date
  'piUnidadeRetorno => 1 = Dia, 2 = Mês, 3 = Ano
  If     piUnidadeRetorno = 1 Then
    ObterDataRetorno = DateAdd("d",piTempoRetorno, pdDataAplicacao)
  ElseIf piUnidadeRetorno = 2 Then
    ObterDataRetorno = DateAdd("m",piTempoRetorno, pdDataAplicacao)
  ElseIf piUnidadeRetorno = 3 Then
    ObterDataRetorno = DateAdd("yyyy",piTempoRetorno, pdDataAplicacao)
  End If
End Function

Public Function InserirRegistroRetorno(ByVal piHandleAplicacao, ByVal piMatricula As Long, ByVal piVacina As Long, ByVal piDose As Integer, pdDataRetorno As Date) As Boolean
'Só será inserido o registro se o mesmo já não existir, pois do contrário será alterado aquele existente
  Dim SQL As Object
  Set SQL = NewQuery
  Dim viHandleExistente As Long

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE                 ")
  SQL.Add("  FROM CLI_VACINA_PACIENTE    ")
  SQL.Add(" WHERE MATRICULA = :MATRICULA ")
  SQL.Add("   AND VACINA    = :VACINA    ")
  SQL.Add("   AND DOSE      = :DOSE      ")
  SQL.Add("   AND HANDLE    > :HANDLE    ")
  SQL.Add("   AND DATAAPLICACAO IS NULL  ")
  SQL.ParamByName("MATRICULA").AsInteger = piMatricula
  SQL.ParamByName("VACINA").AsInteger    = piVacina
  SQL.ParamByName("DOSE").AsInteger      = piDose
  SQL.ParamByName("HANDLE").AsInteger    = piHandleAplicacao
  SQL.Active = True

  If SQL.EOF Then
    SQL.Active = False
    SQL.Clear
    SQL.Add("INSERT INTO CLI_VACINA_PACIENTE  ")
    SQL.Add(" (HANDLE,                        ")
    SQL.Add("  MATRICULA,                     ")
    SQL.Add("  VACINA,                        ")
    SQL.Add("  DOSE,                          ")
    SQL.Add("  USUARIOINCLUSAO,               ")
    SQL.Add("  DATAHORAINCLUSAO,              ")
    SQL.Add("  DATARETORNO)                   ")
    SQL.Add("VALUES                           ")
    SQL.Add(" (:HANDLE,                       ")
    SQL.Add("  :MATRICULA,                    ")
    SQL.Add("  :VACINA,                       ")
    SQL.Add("  :DOSE,                         ")
    SQL.Add("  :USUARIOINCLUSAO,              ")
    SQL.Add("  :DATAHORAINCLUSAO,             ")
    SQL.Add("  :DATARETORNO)                  ")
    SQL.ParamByName("HANDLE").AsInteger            = NewHandle("CLI_VACINA_PACIENTE")
    SQL.ParamByName("MATRICULA").AsInteger         = piMatricula
    SQL.ParamByName("VACINA").AsInteger            = piVacina
    SQL.ParamByName("DOSE").AsInteger              = piDose
    SQL.ParamByName("USUARIOINCLUSAO").AsInteger   = CurrentUser
    SQL.ParamByName("DATAHORAINCLUSAO").AsDateTime = ServerNow
    SQL.ParamByName("DATARETORNO").AsDateTime      = pdDataRetorno
    SQL.ExecSQL
  Else
    viHandleExistente = SQL.FieldByName("HANDLE").AsInteger
    SQL.Active = False
    SQL.Clear
    SQL.Add("UPDATE CLI_VACINA_PACIENTE        ")
    SQL.Add("   SET DATARETORNO = :DATARETORNO ")
    SQL.Add(" WHERE HANDLE = :HANDLEEXISTENTE  ")
    SQL.ParamByName("HANDLEEXISTENTE").AsInteger       = viHandleExistente
    SQL.ParamByName("DATARETORNO").AsDateTime          = pdDataRetorno
    SQL.ExecSQL
  End If

  InserirRegistroRetorno = True
  SQL.Active = False
  Set SQL = Nothing
End Function
