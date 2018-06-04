'HASH: 3CC7DB096BBECBF5C7DB789C15CFFF4A
'#Uses "*bsShowMessage"

Public Sub BOTAORECUSAR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery

  If Not InTransaction Then StartTransaction

  SQL.Add("UPDATE SAM_ROTINAIMP_BENEF SET SITUACAO = 'R' WHERE IMPFAM =:IMPFAM")
  SQL.ParamByName("IMPFAM").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL

  SQL.Clear
  SQL.Add("UPDATE SAM_ROTINAIMP_FAM SET FAMILIAREJEITADA = 'S' WHERE HANDLE =:HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL

  If InTransaction Then Commit

  Set SQL = Nothing

  RefreshNodesWithTable("SAM_ROTINAIMP_FAM")

End Sub

Public Sub BOTAOREPROCESSARERROS_OnClick()
  Dim IMPORTA As Object
  Dim qRotina As Object
  Set qRotina = NewQuery
  Dim vsRetornoMensagem As Long

  qRotina.Active = False
  qRotina.Clear
  qRotina.Add("SELECT R.TABTIPOIMPORTACAO")
  qRotina.Add("  FROM SAM_ROTINAIMP R, SAM_ROTINAIMP_FILIAL F")
  qRotina.Add(" WHERE R.HANDLE = F.ROTINAIMP AND F.HANDLE = :HROTINAFILIAL")
  qRotina.ParamByName("HROTINAFILIAL").AsInteger = CurrentQuery.FieldByName("ROTINAIMPFILIAL").AsInteger
  qRotina.Active = True

  If qRotina.FieldByName("TABTIPOIMPORTACAO").AsInteger = 3 Then

    If VisibleMode Then
      Set IMPORTA = CreateBennerObject("BSINTERFACE0015.RotinasImportacaoBenef")
      vsRetornoMensagem = IMPORTA.ReprocessarErros(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ROTINAIMPFILIAL").AsInteger) 'Handle da filial
    Else
      Set IMPORTA = CreateBennerObject("BSBEN015.ImportarReprocessarErros")
      vsRetornoMensagem = IMPORTA.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ROTINAIMPFILIAL").AsInteger) 'Handle da filial
    End If


    If vsRetornoMensagem = 1 Then
      bsShowMessage("Ocorreu erro no processo","I")
    End If

  Else
    'Set IMPORTA = CreateBennerObject("BSBEN005.Rotinas")
    'IMPORTA.ReprocessarErros(CurrentSystem, CurrentQuery.FieldByName("ROTINAIMPFILIAL").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)

    If VisibleMode Then
      Set IMPORTA = CreateBennerObject("BSINTERFACE0025.RotinasImportacaoBenef")
      vsRetornoMensagem = IMPORTA.ReprocessarErros(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ROTINAIMPFILIAL").AsInteger)
     Else
       Set IMPORTA = CreateBennerObject("BSBEN005.RotinaImportar_ReprocessarErros")
       vsRetornoMensagem = IMPORTA.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ROTINAIMPFILIAL").AsInteger)
    End If

    If vsRetornoMensagem = 1 Then
      bsShowMessage("Ocorreu erro no processo","I")
    End If

  End If

  Set IMPORTA = Nothing
End Sub



Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear

  SQL.Add("SELECT B.SITUACAO, F.ERRO, F.SITUACAO SITUACAOFAM          ")
  SQL.Add("  FROM SAM_ROTINAIMP_FAM        F                          ")
  SQL.Add("  LEFT JOIN SAM_ROTINAIMP_BENEF B ON (B.IMPFAM = F.HANDLE AND B.SITUACAO = 'E'  ) ")
  SQL.Add(" WHERE F.HANDLE = :HANDLE                                  ")

  SQL.Text = SqlConverte(SQL.Text, SQLServer)
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If SQL.FieldByName("SITUACAO").AsString <>"E" And SQL.FieldByName("ERRO").AsString <>"S" And _
     (SQL.FieldByName("SITUACAO").AsString <>"D" And SQL.FieldByName("SITUACAOFAM").AsString <>"D") Then
    BOTAOREPROCESSARERROS.Enabled = False
  Else
    BOTAOREPROCESSARERROS.Enabled = True
  End If

  If SQL.FieldByName("SITUACAO").AsString = "E" Or SQL.FieldByName("SITUACAO").AsString = "G" Or _
	 SQL.FieldByName("ERRO").AsString = "S" Then
    BOTAORECUSAR.Enabled = False
  Else
    BOTAORECUSAR.Enabled = True
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
    Case "BOTAORECUSAR"
      BOTAORECUSAR_OnClick
    Case "BOTAOREPROCESSARERROS"
      BOTAOREPROCESSARERROS_OnClick
  End Select
End Sub
