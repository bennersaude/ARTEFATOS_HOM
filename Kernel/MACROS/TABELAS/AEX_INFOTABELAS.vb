'HASH: F7FC0BB43BD53983E9A49F1E0FCCB725
'# Uses "*bsShowMessage"

Public Sub TABELA_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vCabecs As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String
  Dim vTitulo As String

  If TABELA.PopupCase <> 0 Then
    ShowPopup = False
    Set interface = CreateBennerObject("Procura.Procurar")

    vCabecs = "Nome da tabela"
    vCriterio = "NOME LIKE 'AEX_%' "
    vCriterio = vCriterio + "AND HANDLE NOT IN (SELECT TABELA FROM AEX_INFOTABELAS)"
    vColunas = "NOME"
    vTabela = "Z_TABELAS"
    vTitulo = "Tabelas do autorizador off-line"

    vHandle = interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

    If vHandle <> 0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("TABELA").AsInteger = vHandle
    End If
    Set interface = Nothing
  Else
    ShowPopup = True
  End If

  Dim vNome As String
  'Dim vHandle  As Long
  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Add("SELECT NOME FROM Z_TABELAS WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TABELA").AsInteger
  SQL.Active = True

  vNome = SQL.FieldByName("NOME").AsString
  SQL.Active = False

  CurrentQuery.FieldByName("NOME").Value = vNome
  SQL.Clear
  SQL.Add("SELECT COUNT(1) CONTA FROM AEX_INFOTABELAS WHERE TABELA = :TABELA")
  SQL.ParamByName("TABELA").Value = CurrentQuery.FieldByName("TABELA").AsInteger
  SQL.Active = True
  If SQL.FieldByName("CONTA").AsInteger > 1 Then
    bsShowMessage("Esta tabela já esta cadastrada!", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim QVerifica As Object
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser

  Set QVerifica = NewQuery
  QVerifica.Clear
  QVerifica.Add("SELECT COUNT(1) QTD")
  QVerifica.Add("  FROM AEX_INFOTABELAS")
  QVerifica.Add(" WHERE SIGLATABELA = :SIGLA")
  QVerifica.Add("   AND HANDLE <> :HANDLE")
  QVerifica.ParamByName("SIGLA").AsString = CurrentQuery.FieldByName("SIGLATABELA").AsString
  QVerifica.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QVerifica.Active = True

  If QVerifica.FieldByName("QTD").AsInteger > 1 Then
    bsShowMessage("A sigla da tabela já existe.","E")
    CanContinue = False
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT COUNT(1) CONTA FROM AEX_INFOTABELAS WHERE TABELA = :TABELA")
  SQL.ParamByName("TABELA").Value = CurrentQuery.FieldByName("TABELA").AsInteger
  SQL.Active = True

  If SQL.FieldByName("CONTA").AsInteger > 1 Then
    bsShowMessage("Esta tabela já esta cadastrada!","E")
    CanContinue = False
    Exit Sub
  End If

  SQL.Active = False

  Set QVerifica = Nothing

End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
End Sub


