'HASH: A0F49B994535A9F97CD7827993E8AC18
'CLI_ROTALTERARECURSOAGENDACONF


Public Sub BOTAOCONFIRMAR_OnClick()
  Dim DADOS As Object
  Set DADOS = NewQuery
  Dim SQL As Object
  Set SQL = NewQuery
  If Not CurrentQuery.FieldByName("USUARIO").IsNull Then
    MsgBox("Operação inválida!")
    Exit Sub
  Else
    If Not InTransaction Then StartTransaction

    CurrentQuery.Edit
    CurrentQuery.FieldByName("SITUACAO").AsString = "C"
    CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
    CurrentQuery.FieldByName("DATA").AsDateTime = ServerNow
    CurrentQuery.Post

    'Verifica se já foi alterado
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT USUARIOALTEROU FROM CLI_ROTALTERARECURSO WHERE HANDLE = :ROTINA")
    SQL.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("ROTALTERARECURSO").AsInteger
    SQL.Active = True

    If Not SQL.FieldByName("USUARIOALTEROU").IsNull Then
      DADOS.Clear
      DADOS.Add("SELECT R.RECURSOSUBSTITUTO, E.ESPECIALIDADE, RA.DATAHORAREMARCADA")
      DADOS.Add("  FROM CLI_ROTALTERARECURSO R, CLI_ROTALTERARECURSOAGENDACONF RA, CLI_ESCALA E")
      DADOS.Add(" WHERE R.ESCALASUBSTITUTO = E.HANDLE")
      DADOS.Add("   AND RA.ROTALTERARECURSO = R.HANDLE")
      DADOS.Add("   AND RA.HANDLE = :ROTINA")
      DADOS.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      DADOS.Active = True

      Dim BSCLI001DLL As Object
      Set BSCLI001DLL = CreateBennerObject("BSCLI001.ROTINAS")
      BSCLI001DLL.ConfirmaAgenda(CurrentSystem, _
                                 CurrentQuery.FieldByName("PACIENTE").AsInteger, _
                                 DADOS.FieldByName("RECURSOSUBSTITUTO").AsInteger, _
                                 DADOS.FieldByName("ESPECIALIDADE").AsInteger, _
                                 (DADOS.FieldByName("DATAHORAREMARCADA").AsDateTime), _
                                 "C")
      Set BSCLI001DLL = Nothing
    End If

    If InTransaction Then Commit
    RefreshNodesWithTable("CLI_ROTALTERARECURSOAGENDACONF")
  End If
  Set SQL = Nothing
  Set DADOS = Nothing
End Sub

Public Sub BOTAODESMARCAR_OnClick()
  Dim DADOS As Object
  Set DADOS = NewQuery
  Dim SQL As Object
  Set SQL = NewQuery
  If Not CurrentQuery.FieldByName("USUARIO").IsNull Then
    MsgBox("Operação inválida!")
    Exit Sub
  Else
    CurrentQuery.Edit
    CurrentQuery.FieldByName("SITUACAO").AsString = "D"
    CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
    CurrentQuery.FieldByName("DATA").AsDateTime = ServerNow
    CurrentQuery.Post

    'Verifica se já foi alterado
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT USUARIOALTEROU FROM CLI_ROTALTERARECURSO WHERE HANDLE = :ROTINA")
    SQL.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("ROTALTERARECURSO").AsInteger
    SQL.Active = True

    'Se encontrou,foi alterado então deve desmarcar a consulta
    If Not SQL.FieldByName("USUARIOALTEROU").IsNull Then
      DADOS.Clear
      DADOS.Add("SELECT R.RECURSOSUBSTITUTO, E.ESPECIALIDADE, RA.DATAHORAREMARCADA")
      DADOS.Add("  FROM CLI_ROTALTERARECURSO R, CLI_ROTALTERARECURSOAGENDACONF RA, CLI_ESCALA E")
      DADOS.Add(" WHERE R.ESCALASUBSTITUTO = E.HANDLE")
      DADOS.Add("   AND RA.ROTALTERARECURSO = R.HANDLE")
      DADOS.Add("   AND RA.HANDLE = :ROTINA")
      DADOS.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      DADOS.Active = True

      Dim BSCLI001DLL As Object
      Set BSCLI001DLL = CreateBennerObject("BSCLI001.ROTINAS")
      BSCLI001DLL.DesmarcaAgenda(CurrentSystem, _
                                 CurrentQuery.FieldByName("PACIENTE").AsInteger, _
                                 DADOS.FieldByName("RECURSOSUBSTITUTO").AsInteger, _
                                 (DADOS.FieldByName("DATAHORAREMARCADA").AsDateTime))
      Set BSCLI001DLL = Nothing
    End If
  End If

  RefreshNodesWithTable("CLI_ROTALTERARECURSOAGENDACONF")
  Set SQL = Nothing
  Set DADOS = Nothing
End Sub

Public Sub TABLE_AfterScroll()
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

  If Not qAgenda.FieldByName("BENEFICIARIO").IsNull Then
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
  End If

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

