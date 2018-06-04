'HASH: 36D77E41AEE38CE89BB05535AC6BCFEC
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("DATAINICIAL").AsDateTime = ServerDate
  CurrentQuery.FieldByName("DATAFINAL").AsDateTime = ServerDate
  CurrentQuery.FieldByName("ARQUIVO").AsString = "C:\"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    bsShowMessage("A data inicial não pode ser maior que a final", "E")
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.FieldByName("ARQUIVO").IsNull Then
    bsShowMessage("Informar o caminho do arquivo", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim vsMensagem As String
  Dim viRetorno  As Integer

  If VisibleMode Then
    On Error GoTo Erro
    Dim vDLL As Object
    Set vDLL = CreateBennerObject("BSPREVINNE.GerarArquivoEstratificacao")
    viRetorno = vDLL.Exec(CurrentSystem, _
                          CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
                          CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
                          CurrentQuery.FieldByName("ARQUIVO").AsString, _
                          vsMensagem)

    If viRetorno > 0 Then
      If Trim(vsMensagem) <> "" Then
        bsShowMessage(vsMensagem, "E")
      End If
      CanContinue = False
    Else
      If Trim(vsMensagem) <> "" Then
        bsShowMessage(vsMensagem, "I")
      End If
    End If

    GoTo Fim

    Erro:
    bsShowMessage(Err.Description, "E")
    CanContinue = False

    Fim:
    Set vDLL = Nothing
  Else
    Dim vcContainer As CSDContainer
    Dim Obj         As Object

    Set vcContainer = NewContainer

    On Error GoTo xerro
    vcContainer.AddFields("HANDLE:INTEGER;DATAINICIAL:DATETIME;DATAFINAL:DATETIME;ARQUIVO:STRING;")
    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger = 0
    vcContainer.Field("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
    vcContainer.Field("DATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    vcContainer.Field("ARQUIVO").AsString = CurrentQuery.FieldByName("ARQUIVO").AsString

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")


    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                           "BsPrevinne", _
                                           "GerarArquivoEstratificacao", _
                                           "Geração do arquivo de estratificação PREVINNE", _
                                           0, _
                                           "", _
                                           "", _
                                           "", _
                                           "", _
                                           "P", _
                                           False, _
                                           vsMensagem, _
                                           vcContainer)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
    End If

    GoTo xfim

    xerro:
    bsShowMessage(Err.Description + " - " + vsMensagem, "I")

    xfim:
    Set Obj = Nothing
  End If
End Sub

