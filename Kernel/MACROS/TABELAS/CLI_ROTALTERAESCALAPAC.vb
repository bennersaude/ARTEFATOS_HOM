'HASH: 044601E19FB78A9E2C21CAF16A1B664D

'MACRO CLI_ROTALTERAESCALAPAC

Public Sub BOTAOCONFIRMAR_OnClick()
  Dim sql As Object
  Set sql = NewQuery
  If Not InTransaction Then StartTransaction
  sql.Clear
  sql.Add("UPDATE CLI_ROTALTERAESCALAPAC SET CONFIRMADO = 'S', USUARIOCONFIRMADO = :USUARIO, DATACONFIRMADO = :DATA")
  sql.Add("WHERE HANDLE = :HANDLE")
  sql.ParamByName("USUARIO").AsInteger = CurrentUser
  sql.ParamByName("DATA").AsDateTime = ServerNow
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL
  If InTransaction Then Commit
  RefreshNodesWithTable("CLI_ROTALTERAESCALAPAC")
  Set sql = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim vRot1 As String
  Dim vRot2 As String
  Dim vRot3 As String
  Dim vrot4 As String
  Dim QEndereco As Object
  Set QEndereco = NewQuery

  vRot1 = ""
  vRot2 = ""
  vRot3 = ""
  vrot4 = ""

  If Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull Then
    QEndereco.Active = False
    QEndereco.Clear
    QEndereco.Add("Select E.BAIRRO,")
    QEndereco.Add("       E.LOGRADOURO,")
    QEndereco.Add("       E.NUMERO,")
    QEndereco.Add("       E.COMPLEMENTO,")
    QEndereco.Add("       E.TELEFONE1,")
    QEndereco.Add("       E.TELEFONE2,")
    QEndereco.Add("       E.CEP,")
    QEndereco.Add("       T.NOME ESTADO,")
    QEndereco.Add("       M.NOME MUNICIPIO")
    QEndereco.Add("  FROM SAM_ENDERECO E,")
    QEndereco.Add("       SAM_BENEFICIARIO B,")
    QEndereco.Add("       ESTADOS T,")
    QEndereco.Add("       MUNICIPIOS M")
    QEndereco.Add(" WHERE B.ENDERECORESIDENCIAL = E.HANDLE")
    QEndereco.Add("   And E.MUNICIPIO = M.HANDLE")
    QEndereco.Add("   And E.ESTADO = T.HANDLE")
    QEndereco.Add("   And B.HANDLE = :BENEFICIARIO")
    QEndereco.ParamByName("BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    QEndereco.Active = True

    If Not QEndereco.FieldByName("LOGRADOURO").IsNull Then
      vRot1 = "Logradouro: " + QEndereco.FieldByName("LOGRADOURO").AsString + "   "
    End If

    If Not QEndereco.FieldByName("NUMERO").IsNull Then
      vRot1 = vRot1 + "Número: " + QEndereco.FieldByName("NUMERO").AsString
    End If

    If Not QEndereco.FieldByName("COMPLEMENTO").IsNull Then
      vRot2 = "Complemento: " + QEndereco.FieldByName("COMPLEMENTO").AsString + "   "
    End If

    If Not QEndereco.FieldByName("BAIRRO").IsNull Then
      vRot2 = vRot2 + "Bairro: " + QEndereco.FieldByName("BAIRRO").AsString
    End If

    If Not QEndereco.FieldByName("CEP").IsNull Then
      vRot3 = "CEP: " + QEndereco.FieldByName("CEP").AsString + "   "
    End If

    If Not QEndereco.FieldByName("ESTADO").IsNull Then
      vRot3 = vRot3 + "Estado: " + QEndereco.FieldByName("ESTADO").AsString + "   "
    End If

    If Not QEndereco.FieldByName("MUNICIPIO").IsNull Then
      vRot3 = vRot3 + "Município: " + QEndereco.FieldByName("MUNICIPIO").AsString
    End If

    If Not QEndereco.FieldByName("TELEFONE1").IsNull Then
      vrot4 = "Telefone 1: " + QEndereco.FieldByName("TELEFONE1").AsString + "   "
    End If

    If Not QEndereco.FieldByName("TELEFONE2").IsNull Then
      vrot4 = vrot4 + "Telefone 2: " + QEndereco.FieldByName("TELEFONE2").AsString
    End If

  End If

  ROTULOEND1.Text = vRot1
  ROTULOEND2.Text = vRot2
  ROTULOEND3.Text = vRot3
  ROTULOEND4.Text = vrot4


  Set QEndereco = Nothing
End Sub

