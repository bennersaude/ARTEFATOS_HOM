'HASH: 917AF8DF5977EE499FE543A785D7C7A1


Public Sub BOTAOREMARCAR_OnClick()
  If CurrentQuery.FieldByName("USUARIO").IsNull Then
    Dim BSCli001dll As Object
    Set BSCli001dll = CreateBennerObject("BSCli001.Rotinas")
    BSCli001dll.RemarcaAgendaPonte(CurrentSystem, _
                                   CurrentQuery.FieldByName("PACIENTE").AsInteger, _
                                   CurrentQuery.FieldByName("HANDLE").AsInteger)
    RefreshNodesWithTable("CLI_PONTESRECURSOAGENDA")
    Set BSCli001dll = Nothing
  Else
    MsgBox("A consulta já foi remarcada!")
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("USUARIO").IsNull Then
    BOTAOREMARCAR.Enabled = True
  Else
    BOTAOREMARCAR.Enabled = False
  End If

  Dim vRot1 As String
  Dim vRot2 As String
  Dim vRot3 As String
  Dim vrot4 As String
  Dim vrot5 As String
  Dim QEndereco As Object
  Set QEndereco = NewQuery
  Dim qAgenda As Object
  Set qAgenda = NewQuery

  vRot1 = ""
  vRot2 = ""
  vRot3 = ""
  vrot4 = ""
  vrot5 = ""

  qAgenda.Active = False
  qAgenda.Clear
  qAgenda.Add("SELECT MATRICULA, BENEFICIARIO, ENDERECO, TELEFONECONTATO FROM CLI_AGENDA")
  qAgenda.Add(" WHERE HANDLE = :AGENDA")
  qAgenda.ParamByName("AGENDA").AsInteger = CurrentQuery.FieldByName("PACIENTE").AsInteger
  qAgenda.Active = True

  If Not qAgenda.FieldByName("ENDERECO").IsNull Then
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
    QEndereco.Add("       ESTADOS T,")
    QEndereco.Add("       MUNICIPIOS M")
    QEndereco.Add(" WHERE E.MUNICIPIO = M.HANDLE")
    QEndereco.Add("   And E.ESTADO = T.HANDLE")
    QEndereco.Add("   And E.HANDLE = :ENDERECO")
    QEndereco.ParamByName("ENDERECO").AsInteger = qAgenda.FieldByName("ENDERECO").AsInteger
    QEndereco.Active = True

  ElseIf Not qAgenda.FieldByName("BENEFICIARIO").IsNull Then
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
    QEndereco.ParamByName("BENEFICIARIO").AsInteger = qAgenda.FieldByName("BENEFICIARIO").AsInteger
    QEndereco.Active = True
  End If

  If Not QEndereco.EOF Then
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

  If Not qAgenda.FieldByName("TELEFONECONTATO").IsNull Then
    vrot5 = "Telefone contato: " + qAgenda.FieldByName("TELEFONECONTATO").AsString
  End If


  ROTULOEND1.Text = vRot1
  ROTULOEND2.Text = vRot2
  ROTULOEND3.Text = vRot3
  ROTULOEND4.Text = vrot4
  ROTULOEND5.Text = vrot5


  Set QEndereco = Nothing
  Set qAgenda = Nothing
End Sub

