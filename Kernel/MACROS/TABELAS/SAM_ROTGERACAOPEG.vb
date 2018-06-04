'HASH: 84BFAE79802CDC3A29CDAD3BAB18BC00
'SAM_ROTGERACAOPEG
'Alteração: Milton/17/01/2002 -SMS 5976
'Última alteração: Daniela 22/04/2002 -SMS 7189


Public Sub BOTAOPROCESSAR_OnClick()
  Dim interface As Object
  If CurrentQuery.State <>1 Then
    MsgBox("O registro está em edição! Por favor confirme ou cancele as alterações")
    Exit Sub
  End If
  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    MsgBox("Rotina já processada!")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABROTINA").AsInteger = 1 Then
    Set interface = CreateBennerObject("CliGeraGuia.GeraGuia")
    If CurrentQuery.FieldByName("TipoAtendimento").Value = "M" Then 'Médico =GS
      interface.Gerar(CurrentSystem, "1", _
                      CurrentQuery.FieldByName("competencia").AsDateTime, _
                      CurrentQuery.FieldByName("CartaRemessa").AsInteger, _
                      CurrentQuery.FieldByName("PEG").AsInteger, _
                      CurrentQuery.FieldByName("CartaRemessaExame").AsInteger, _
                      CurrentQuery.FieldByName("PEGExame").AsInteger, _
                      CurrentQuery.FieldByName("dataRecebimento").AsDateTime, _
                      CurrentQuery.FieldByName("Clinica").AsInteger, _
                      CurrentQuery.FieldByName("handle").AsInteger)
    Else 'Coletivo =GC
      interface.Gerar(CurrentSystem, "2", _
                      CurrentQuery.FieldByName("competencia").AsDateTime, _
                      CurrentQuery.FieldByName("CartaRemessa").AsInteger, _
                      CurrentQuery.FieldByName("PEG").AsInteger, _
                      CurrentQuery.FieldByName("CartaRemessaExame").AsInteger, _
                      CurrentQuery.FieldByName("PEGExame").AsInteger, _
                      CurrentQuery.FieldByName("dataRecebimento").AsDateTime, _
                      CurrentQuery.FieldByName("Clinica").AsInteger, _
                      CurrentQuery.FieldByName("handle").AsInteger)
    End If
  Else 'TabRotina =2
    Dim MyQuery As Object
    Set interface = CreateBennerObject("SAMBDAM.BOLETIM")
    Set MyQuery = NewQuery
    MyQuery.Active = False
    MyQuery.Clear
    MyQuery.Add("Select prestador from at_Clinica where Handle= :Clinica")
    MyQuery.ParamByName("Clinica").Value = CurrentQuery.FieldByName("CLINICABDAM").Value
    MyQuery.Active = True

    If CurrentQuery.FieldByName("TipoAtendimento").Value = "M" Then
      interface.Executar(CurrentSystem, "1", _
                         CurrentQuery.FieldByName("competencia").AsDateTime, _
                         CurrentQuery.FieldByName("CartaRemessa").Value, _
                         CurrentQuery.FieldByName("PEG").Value, _
                         CurrentQuery.FieldByName("dataRecebimento").AsDateTime, _
                         CurrentQuery.FieldByName("ClinicaBdam").Value, _
                         MyQuery.FieldByName("prestador").Value, _
                         CurrentQuery.FieldByName("handle").Value)
    Else
      interface.Executar(CurrentSystem, "2", _
                         CurrentQuery.FieldByName("competencia").AsDateTime, _
                         CurrentQuery.FieldByName("CartaRemessa").Value, _
                         CurrentQuery.FieldByName("PEG").Value, _
                         CurrentQuery.FieldByName("dataRecebimento").AsDateTime, _
                         CurrentQuery.FieldByName("ClinicaBdam").Value, _
                         MyQuery.FieldByName("prestador").Value, _
                         CurrentQuery.FieldByName("handle").Value)
    End If
    Set MyQuery = Nothing
  End If

  RefreshNodesWithTable("SAM_ROTGERACAOPEG")
  Set interface = Nothing

End Sub


Public Sub CARTAREMESSA_OnExit()
  If Not CurrentQuery.FieldByName("PEG").IsNull Then
    Exit Sub
  End If

  Dim VALOR As Long
  Dim NUMERACAO As Object
  Set NUMERACAO = NewQuery
  NUMERACAO.Active = False
  NUMERACAO.Clear
  NUMERACAO.Add("SELECT NUMERACAOPEG FROM SAM_PARAMETROSPROCCONTAS")
  NUMERACAO.Active = True

  If Not CurrentQuery.FieldByName("CARTAREMESSA").IsNull Then
    If NUMERACAO.FieldByName("NUMERACAOPEG").AsString = "M" Then
      CurrentQuery.FieldByName("PEG").AsInteger = CurrentQuery.FieldByName("CARTAREMESSA").AsInteger
    End If
  End If
  Set NUMERACAO = Nothing

End Sub

Public Sub CLINICA_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PRESTADOR.NOME|SAM_PRESTADOR.PRESTADOR"

  vCampos = "Nome|CNPJ/CPF"

  vHandle = interface.Exec(CurrentSystem, "CLI_CLINICA|SAM_PRESTADOR[SAM_PRESTADOR.HANDLE=CLI_CLINICA.PRESTADOR]", vColunas, 1, vCampos, vCriterio, "Clínicas Próprias", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CLINICA").Value = vHandle
  End If
  Set interface = Nothing

End Sub


Public Sub PEG_OnExit()

  If Not CurrentQuery.FieldByName("competencia").IsNull Then
    If CurrentQuery.State = 3 Then 'inserir
      Dim q1 As Object
      Dim vdata As Date
      Set q1 = NewQuery
      Set q2 = NewQuery
      'Encontrar handle referente a competência informada na Rotina
      q2.Clear
      q2.Add("SELECT HANDLE FROM SAM_COMPETPEG WHERE COMPETENCIA=:COMPETENCIA")
      vdata = CDate(Format(CurrentQuery.FieldByName("competencia").AsString, "yyyy") + "-01-01")
      q2.ParamByName("competencia").Value = vdata
      q2.Active = True
      q1.Clear
      q1.Add("SELECT HANDLE FROM SAM_PEG WHERE COMPETENCIA=:COMPETENCIA AND PEG=:PEG")
      q1.ParamByName("COMPETENCIA").Value = q2.FieldByName("handle").Value
      q1.ParamByName("PEG").Value = CurrentQuery.FieldByName("PEG").AsInteger
      q1.Active = True
      If Not q1.EOF Then
        MsgBox("Já existe um PEG com este número")
        PEG.SetFocus
      End If
      q1.Active = False
      q2.Active = False
      Set q1 = Nothing
      Set q2 = Nothing
    End If
  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  If CurrentQuery.FieldByName("situacao").Value = "P" Then
    MsgBox("Impossível excluir registro porque rotina já foi processada!")
    CanContinue = False
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If CurrentQuery.FieldByName("situacao").Value = "P" Then
    MsgBox("Impossível alterar campo porque rotina já foi processada!")
    CanContinue = False
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("TIPOATENDIMENTO").AsString = "M" Then
    If CurrentQuery.FieldByName("TABROTINA").AsInteger = 1 Then
      If CurrentQuery.FieldByName("PEGEXAME").IsNull Then
        MsgBox("Para atendimento médico o número do PEG de exames é obrigatório!")
        CanContinue = False
        Exit Sub
      Else
        If CurrentQuery.FieldByName("PEGEXAME").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger Then
          MsgBox("O número do PEG de exame não pode ser igual ao PEG de serviço!")
          CanContinue = False
          Exit Sub
        End If
      End If
      If CurrentQuery.FieldByName("CARTAREMESSAEXAME").IsNull Then
        MsgBox("Para atendimento médico o número da carta remessa de exames é obrigatório!")
        CanContinue = False
        Exit Sub
      End If
    End If
  End If
End Sub

