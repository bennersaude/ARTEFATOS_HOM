'HASH: 7A902FB8AA71BC7C0BD8DDF0B0AB0B92
'#Uses "*bsShowMessage"
'CLI_AGENDA


Public Sub BOTAOCHEGADAPACIENTE_OnClick()

  If CurrentQuery.State = 1 Then
    If CurrentQuery.FieldByName("DATACHEGADA").IsNull Then
      Dim sql As Object
      Set sql = NewQuery
      sql.Clear
      sql.Add("SELECT HANDLE FROM CLI_AGENDA_NEGACAO WHERE AGENDA = :AGENDA AND SITUACAO = 'P'")
      sql.ParamByName("AGENDA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      sql.Active = True
      If Not sql.EOF Then
        bsShowMessage("Existem negações pendentes para esse beneficiário!","I")
        Exit Sub
      End If

      If CurrentQuery.FieldByName("DATAMARCADA").AsDateTime <>ServerDate Then
        bsShowMessage("Não é possível marcar chegada para uma consulta de outro dia!", "I")
        Exit Sub
      End If

      If Not InTransaction Then StartTransaction

      sql.Clear
      sql.Add("UPDATE CLI_AGENDA SET DATACHEGADA = :DATA, HORACHEGADA = :HORA, USUARIOCHEGADA = :USERCHEGADA WHERE HANDLE = :HANDLE")
      sql.ParamByName("DATA").Value = ServerDate
      sql.ParamByName("HORA").Value = TimeValue(Str(DatePart("h", ServerNow)) + ":" + Str(DatePart("n", ServerNow)))
      sql.ParamByName("USERCHEGADA").AsInteger = CurrentUser
      sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      sql.ExecSQL

      If InTransaction Then Commit

      RefreshNodesWithTable("CLI_AGENDA")
      Set sql = Nothing
    End If
  Else
    bsShowMessage("A tabela não pode estar em edição ou inserção!", "I")
  End If

End Sub


Public Sub BOTAOCONFIRMAR_OnClick()
  If Not (CurrentQuery.FieldByName("HORACHEGADA").IsNull) Then
    bsShowMessage("Confirmação não realizada, pois há registro de chegada do paciente!", "I")
    Exit Sub
  Else
    If Not (CurrentQuery.FieldByName("DATAHORACONFIRMACAO").IsNull) Then
      bsShowMessage("Atendimento já confirmado!", "I")
      Exit Sub
    Else
      If Not InTransaction Then StartTransaction

      Dim qConfirmaAtendimento As Object
      Set qConfirmaAtendimento = NewQuery
      qConfirmaAtendimento.Clear
      qConfirmaAtendimento.Add("UPDATE CLI_AGENDA SET DATAHORACONFIRMACAO = :DHCONFIRMACAO,   ")
      qConfirmaAtendimento.Add("                       USUARIOCONFIRMACAO = :USERCONFIRMACAO  ")
      qConfirmaAtendimento.Add("                             WHERE HANDLE = :HANDLE           ")
      qConfirmaAtendimento.ParamByName("DHCONFIRMACAO").AsDateTime = ServerNow
      qConfirmaAtendimento.ParamByName("USERCONFIRMACAO").AsInteger = CurrentUser
      qConfirmaAtendimento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qConfirmaAtendimento.ExecSQL

      If InTransaction Then Commit

      RefreshNodesWithTable("CLI_AGENDA")
      Set qConfirmaAtendimento = Nothing

    End If
  End If
End Sub

Public Sub BOTAOLIMPARCONFIRMACAO_OnClick()
    If Not (CurrentQuery.FieldByName("HORACHEGADA").IsNull) Then
      bsShowMessage("Confirmação não pode ser apagada, pois há registro de chegada do paciente!", "I")
      Exit Sub
  Else
    If CurrentQuery.FieldByName("DATAHORACONFIRMACAO").IsNull Then
      bsShowMessage("Atendimento ainda não confirmado!", "I")
      Exit Sub
    Else
      If Not InTransaction Then StartTransaction

      Dim qConfirmaAtendimento As Object
      Set qConfirmaAtendimento = NewQuery
      qConfirmaAtendimento.Clear
      qConfirmaAtendimento.Add("UPDATE CLI_AGENDA                   ")
      qConfirmaAtendimento.Add("   SET DATAHORACONFIRMACAO = NULL,  ")
      qConfirmaAtendimento.Add("       USUARIOCONFIRMACAO = NULL    ")
      qConfirmaAtendimento.Add(" WHERE HANDLE = :HANDLE             ")
      qConfirmaAtendimento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qConfirmaAtendimento.ExecSQL

      If InTransaction Then Commit

      RefreshNodesWithTable("CLI_AGENDA")
      Set qConfirmaAtendimento = Nothing

    End If
  End If
End Sub

Public Sub BOTAOREVERTER_OnClick()
  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("CliClinica.Agenda")
  AGENDA.ReverterNegacao(CurrentSystem, CurrentQuery.FieldByName("EVENTO").AsInteger, _
                         CurrentUser, _
                         CurrentQuery.FieldByName("HANDLE").AsInteger, _
                         "A")
  Set AGENDA = Nothing
End Sub

Public Sub BOTAOREVERTEROUTROUSUARIO_OnClick()
  Dim OLESenha As Object
  Set OLESenha = CreateBennerObject("senha.rotinas")
  vUsuario = OLESenha.PegarUsuarioPadrao(CurrentSystem, CurrentUser, vUsuario)
  Set OLESenha = Nothing
  If vUsuario <0 Then
    Exit Sub
  End If
  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("CliClinica.Agenda")
  AGENDA.ReverterNegacao(CurrentSystem, CurrentQuery.FieldByName("EVENTO").AsInteger, _
                         vUsuario, _
                         CurrentQuery.FieldByName("HANDLE").AsInteger, _
                         "A")
  Set AGENDA = Nothing
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

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim OK As Boolean
  Dim AGENDA As Object
  Set AGENDA = CreateBennerObject("CliClinica.Agenda")

  AGENDA.VerificaEscala(CurrentSystem, CurrentQuery.FieldByName("DATAMARCADA").AsDateTime, _
                        CurrentQuery.FieldByName("RECURSO").AsInteger, _
                        CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, _
                        OK)
  If Not OK Then
    CanContinue = False
    bsShowMessage("O horário escolhido não possui escala!", "E")
  End If
  Set AGENDA = Nothing

  If CurrentQuery.FieldByName("MATRICULA").IsNull Then
    Dim BENEF As Object
    Set BENEF = NewQuery
    BENEF.Add("SELECT MATRICULA FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")
    BENEF.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    BENEF.Active = True
    If Not BENEF.EOF Then
      CurrentQuery.FieldByName("MATRICULA").AsInteger = BENEF.FieldByName("MATRICULA").AsInteger
    End If
    Set BENEF = Nothing
  End If
  If Not CurrentQuery.FieldByName("MOTIVODESMARCACAO").IsNull Then
    If CurrentQuery.FieldByName("DATADESMARCACAO").IsNull Then
      CurrentQuery.FieldByName("DATADESMARCACAO").AsDateTime = ServerDate
      CurrentQuery.FieldByName("HORADESMARCACAO").AsDateTime = TimeValue(Str(DatePart("h", ServerNow)) + ":" + Str(DatePart("n", ServerNow)))
      CurrentQuery.FieldByName("USUARIODESMARQUE").AsInteger = CurrentUser
    End If
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
 If (CommandID = "BOTAOCHEGADAPACIENTE") Then
	BOTAOCHEGADAPACIENTE_OnClick
 End If
 If (CommandID = "BOTAOREVERTER") Then
	BOTAOREVERTER_OnClick
 End If
 If (CommandID = "BOTAOREVERTEROUTROUSUARIO") Then
	BOTAOREVERTEROUTROUSUARIO_OnClick
 End If
End Sub
