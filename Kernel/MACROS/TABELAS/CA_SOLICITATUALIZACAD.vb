'HASH: 6F17B6B896ACD7FD0A21F28142F9449E

'############## CENTRAL DE ATENDIMENTO #################

Public Sub BOTAOCANCELAR_OnClick()
  ' +++++++++dentro da dll da a mensagem de confirmacao
  '	If MsgBox("Confirma o cancelamento da solicitação?",vbYesNo)=vbNo Then
  '       Exit Sub
  '    End If

  Dim vDll As Object
  Dim vRetorno As Boolean

  Set vDll = CreateBennerObject("CA024.SolAlteracaoCad")

  '"CA_SOLICITATUALIZACAD",

  vRetorno = vDll.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  If vRetorno = False Then
    Exit Sub
  End If

  WriteAudit("C", HandleOfTable("CA_SOLICITATUALIZACAD"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Solicitação de alteração cadastral - Cancelamento")


  RefreshNodesWithTable("CA_SOLICITATUALIZACAD")
End Sub

Public Sub TABEND_OnChanging(AllowChange As Boolean)

  'AllowChange =False

End Sub

Public Sub TABLE_AfterScroll()
  Dim ENDERECO As Object
  Dim COMERCIAL As Object
  Dim CORRESPONDENCIA As Object
  Dim ESTADO As Object
  Dim vMunicipio, vEstado, vMunicipioCom, vEstadoCom, vMunicipioCor, vEstadoCor As Long
  Dim vNumero, vComplemento, vBairro, vCEP, vTelefone1, vTelefone2, vFax As String
  Dim vNumeroCom, vComplementoCom, vBairroCom, vCEPCom, vTelefone1Com, vTelefone2Com, vFaxCom As String
  Dim vNumeroCor, vComplementoCor, vBairroCor, vCEPCor, vTelefone1Cor, vTelefone2Cor, vFaxCor As String
  Dim vRot1, vRot2, vRot3, vRot4 As String
  Dim vRot1Com, vRot2Com, vRot3Com, vRot4Com As String
  Dim vRot1Cor, vRot2Cor, vRot3Cor, vRot4Cor As String
  Set ENDERECO = NewQuery
  Set COMERCIAL = NewQuery
  Set CORRESPONDENCIA = NewQuery
  Set ESTADO = NewQuery

  ENDERECO.Active = False
  ENDERECO.Clear
  ENDERECO.Add("SELECT ESTADO, MUNICIPIO, BAIRRO, CEP, NUMERO, COMPLEMENTO, TELEFONE1, TELEFONE2, FAX, LOGRADOURO")
  ENDERECO.Add("  FROM SAM_ENDERECO")
  ENDERECO.Add(" WHERE HANDLE = :ENDERECORESIDENCIAL")
  ENDERECO.ParamByName("ENDERECORESIDENCIAL").Value = CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger
  ENDERECO.Active = True
  vLogradouro = ENDERECO.FieldByName("LOGRADOURO").AsString
  vNumero = ENDERECO.FieldByName("NUMERO").AsString
  vComplemento = ENDERECO.FieldByName("COMPLEMENTO").AsString
  vBairro = ENDERECO.FieldByName("BAIRRO").AsString
  vCEP = ENDERECO.FieldByName("CEP").AsString
  vMunicipio = ENDERECO.FieldByName("MUNICIPIO").AsInteger
  vEstado = ENDERECO.FieldByName("ESTADO").AsInteger
  vTelefone1 = ENDERECO.FieldByName("TELEFONE1").AsString
  vTelefone2 = ENDERECO.FieldByName("TELEFONE2").AsString
  vFax = ENDERECO.FieldByName("FAX").AsString

  ESTADO.Active = False
  ESTADO.Clear
  ESTADO.Add("Select E.NOME NOMEESTADO, M.NOME NOMEMUNICIPIO")
  ESTADO.Add("  FROM ESTADOS E,")
  ESTADO.Add("  MUNICIPIOS M")
  ESTADO.Add(" WHERE M.ESTADO = E.HANDLE")
  ESTADO.Add("   And M.HANDLE = :HANDLEMUNICIPIO")
  ESTADO.Add("   And E.HANDLE = :HANDLEESTADO")
  ESTADO.ParamByName("HANDLEMUNICIPIO").Value = vMunicipio
  ESTADO.ParamByName("HANDLEESTADO").Value = vEstado
  ESTADO.Active = True


  vRot1 = ""
  vRot2 = ""
  vRot3 = ""
  vRot4 = ""

  If vLogradouro <>"" Then
    vRot1 = "Logradouro: " + vLogradouro
  End If

  If vNumero <>"" Then
    vRot1 = vRot1 + " Número: " + vNumero
  End If

  If vComplemento <>"" Then
    vRot2 = "Complemento: " + vComplemento + "     "
  End If

  If vBairro <>"" Then
    vRot2 = vRot2 + "Bairro: " + vBairro
  End If

  If vCEP <>"" Then
    vRot3 = "CEP: " + vCEP + "     "
  End If

  If vMunicipio <>0 Then
    vRot3 = vRot3 + "Município: " + ESTADO.FieldByName("NOMEMUNICIPIO").AsString + "     "
  End If

  If vEstado <>0 Then
    vRot3 = vRot3 + "Estado: " + ESTADO.FieldByName("NOMEESTADO").AsString
  End If

  If vTelefone1 <>"" Then
    vRot4 = "Telefone 1: " + vTelefone1 + "     "
  End If

  If vTelefone2 <>"" Then
    vRot4 = vRot4 + "Telefone 2: " + vTelefone2 + "     "
  End If

  If vFax <>"" Then
    vRot4 = vRot4 + "Fax: " + vFax
  End If

  ROTULORES1.Text = vRot1
  ROTULORES2.Text = vRot2
  ROTULORES3.Text = vRot3
  ROTULORES4.Text = vRot4



  COMERCIAL.Active = False
  COMERCIAL.Clear
  COMERCIAL.Add("SELECT ESTADO, MUNICIPIO, BAIRRO, CEP, NUMERO, COMPLEMENTO, TELEFONE1, TELEFONE2, FAX, LOGRADOURO")
  COMERCIAL.Add("  FROM SAM_ENDERECO")
  COMERCIAL.Add(" WHERE HANDLE = :ENDERECOCOMERCIAL")
  COMERCIAL.ParamByName("ENDERECOCOMERCIAL").Value = CurrentQuery.FieldByName("ENDERECOCOMERCIAL").AsInteger
  COMERCIAL.Active = True
  vLogradourocom = COMERCIAL.FieldByName("LOGRADOURO").AsString
  vNumeroCom = COMERCIAL.FieldByName("NUMERO").AsString
  vComplementoCom = COMERCIAL.FieldByName("COMPLEMENTO").AsString
  vBairroCom = COMERCIAL.FieldByName("BAIRRO").AsString
  vCEPCom = COMERCIAL.FieldByName("CEP").AsString
  vMunicipioCom = COMERCIAL.FieldByName("MUNICIPIO").AsInteger
  vEstadoCom = COMERCIAL.FieldByName("ESTADO").AsInteger
  vTelefone1Com = COMERCIAL.FieldByName("TELEFONE1").AsString
  vTelefone2Com = COMERCIAL.FieldByName("TELEFONE2").AsString
  vFaxCom = COMERCIAL.FieldByName("FAX").AsString

  ESTADO.Active = False
  ESTADO.Clear
  ESTADO.Add("Select E.NOME NOMEESTADO, M.NOME NOMEMUNICIPIO")
  ESTADO.Add("  FROM ESTADOS E,")
  ESTADO.Add("  MUNICIPIOS M")
  ESTADO.Add(" WHERE M.ESTADO = E.HANDLE")
  ESTADO.Add("   And M.HANDLE = :HANDLEMUNICIPIO")
  ESTADO.Add("   And E.HANDLE = :HANDLEESTADO")
  ESTADO.ParamByName("HANDLEMUNICIPIO").Value = vMunicipioCom
  ESTADO.ParamByName("HANDLEESTADO").Value = vEstadoCom
  ESTADO.Active = True

  vRot1Com = ""
  vRot2Com = ""
  vRot3Com = ""
  vRot4Com = ""

  If vLogradourocom <>"" Then
    vRot1Com = "Logradouro: " + vLogradourocom
  End If

  If vNumeroCom <>"" Then
    vRot1Com = vRot1Com + " Número: " + vNumeroCom
  End If

  If vComplementoCom <>"" Then
    vRot2Com = "Complemento: " + vComplementoCom + "     "
  End If

  If vBairroCom <>"" Then
    vRot2Com = vRot2Com + "Bairro: " + vBairroCom
  End If

  If vCEPCom <>"" Then
    vRot3Com = "CEP: " + vCEPCom + "     "
  End If

  If vMunicipioCom <>0 Then
    vRot3Com = vRot3Com + "Município: " + ESTADO.FieldByName("NOMEMUNICIPIO").AsString + "     "
  End If

  If vEstadoCom <>0 Then
    vRot3Com = vRot3Com + "Estado: " + ESTADO.FieldByName("NOMEESTADO").AsString
  End If

  If vTelefone1Com <>"" Then
    vRot4Com = "Telefone 1: " + vTelefone1Com + "     "
  End If

  If vTelefone2Com <>"" Then
    vRot4Com = vRot4Com + "Telefone 2: " + vTelefone2Com + "     "
  End If

  If vFaxCom <>"" Then
    vRot4Com = vRot4Com + "Fax: " + vFaxCom
  End If

  ROTULOCOM1.Text = vRot1Com
  ROTULOCOM2.Text = vRot2Com
  ROTULOCOM3.Text = vRot3Com
  ROTULOCOM4.Text = vRot4Com

  CORRESPONDENCIA.Active = False
  CORRESPONDENCIA.Clear
  CORRESPONDENCIA.Add("SELECT ESTADO, MUNICIPIO, BAIRRO, CEP, NUMERO, COMPLEMENTO, TELEFONE1, TELEFONE2, FAX, LOGRADOURO")
  CORRESPONDENCIA.Add("  FROM SAM_ENDERECO")
  CORRESPONDENCIA.Add(" WHERE HANDLE = :ENDERECOCORRESPONDENCIA")
  CORRESPONDENCIA.ParamByName("ENDERECOCORRESPONDENCIA").Value = CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger
  CORRESPONDENCIA.Active = True

  vLogradouroCor = CORRESPONDENCIA.FieldByName("LOGRADOURO").AsString
  vNumeroCor = CORRESPONDENCIA.FieldByName("NUMERO").AsString
  vComplementoCor = CORRESPONDENCIA.FieldByName("COMPLEMENTO").AsString
  vBairroCor = CORRESPONDENCIA.FieldByName("BAIRRO").AsString
  vCEPCor = CORRESPONDENCIA.FieldByName("CEP").AsString
  vMunicipioCor = CORRESPONDENCIA.FieldByName("MUNICIPIO").AsInteger
  vEstadoCor = CORRESPONDENCIA.FieldByName("ESTADO").AsInteger
  vTelefone1Cor = CORRESPONDENCIA.FieldByName("TELEFONE1").AsString
  vTelefone2Cor = CORRESPONDENCIA.FieldByName("TELEFONE2").AsString
  vFaxCor = CORRESPONDENCIA.FieldByName("FAX").AsString

  ESTADO.Active = False
  ESTADO.Clear
  ESTADO.Add("Select E.NOME NOMEESTADO, M.NOME NOMEMUNICIPIO")
  ESTADO.Add("  FROM ESTADOS E,")
  ESTADO.Add("  MUNICIPIOS M")
  ESTADO.Add(" WHERE M.ESTADO = E.HANDLE")
  ESTADO.Add("   And M.HANDLE = :HANDLEMUNICIPIO")
  ESTADO.Add("   And E.HANDLE = :HANDLEESTADO")
  ESTADO.ParamByName("HANDLEMUNICIPIO").Value = vMunicipioCor
  ESTADO.ParamByName("HANDLEESTADO").Value = vEstadoCor
  ESTADO.Active = True

  vRot1Cor = ""
  vRot2Cor = ""
  vRot3Cor = ""
  vRot4Cor = ""

  If vLogradouroCor <>"" Then
    vRot1Cor = "Logradouro: " + vLogradourocor
  End If

  If vNumeroCor <>"" Then
    vRot1Cor = vRot1Cor + " Número: " + vNumeroCor
  End If

  If vComplementoCor <>"" Then
    vRot2Cor = "Complemento: " + vComplementoCor + "     "
  End If

  If vBairroCor <>"" Then
    vRot2Cor = vRot2Cor + "Bairro: " + vBairroCor
  End If

  If vCEPCor <>"" Then
    vRot3Cor = "CEP: " + vCEPCor + "     "
  End If

  If vMunicipioCor <>0 Then
    vRot3Cor = vRot3Cor + "Município: " + ESTADO.FieldByName("NOMEMUNICIPIO").AsString + "     "
  End If

  If vEstadoCor <>0 Then
    vRot3Cor = vRot3Cor + "Estado: " + ESTADO.FieldByName("NOMEESTADO").AsString
  End If

  If vTelefone1Cor <>"" Then
    vRot4Cor = "Telefone 1: " + vTelefone1Cor + "     "
  End If

  If vTelefone2Cor <>"" Then
    vRot4Cor = vRot4Cor + "Telefone 2: " + vTelefone2Cor + "     "
  End If

  If vFaxCor <>"" Then
    vRot4Cor = vRot4Cor + "Fax: " + vFaxCor
  End If

  ROTULOCOR1.Text = vRot1Cor
  ROTULOCOR2.Text = vRot2Cor
  ROTULOCOR3.Text = vRot3Cor
  ROTULOCOR4.Text = vRot4Cor


  Set ENDERECO = Nothing
  Set COMERCIAL = Nothing
  Set CORRESPONDENCIA = Nothing

  'If vMascaraBeneficiario ="" Then
  '  Dim VBENEFICIARIO As String
  '  Dim VCODIGOFORMATADO As String
  '  Dim interface As Object
  '
  '  VBENEFICIARIO =CurrentQuery.FieldByName("BENEFICIARIO").AsString
  '
  '  Set interface=CreateBennerObject("SamBeneficiario.Cadastro")
  '    interface.Mascara(VBENEFICIARIO,vMascaraBeneficiario,VCODIGOFORMATADO)
  '  Set Interface =Nothing
  '
  'End If

  CurrentQuery.FieldByName("BENEFICIARIO").Mask = vMascaraBeneficiario


  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT TABRESPONSAVEL FROM CA_SOLICITATUALIZACAD WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  'MsgBox(CStr(TABEND.PageIndex))
  'TABEND.PageIndex  =CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger -1

  Select Case CurrentQuery.FieldByName("SITUACAO").AsString
    Case "C"
      BOTAOCANCELAR.Visible = False
    Case "P"
      BOTAOCANCELAR.Visible = False
    Case Else
      BOTAOCANCELAR.Visible = True
  End Select

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  Select Case CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger
    Case 1
      SQL.Clear
      SQL.Add("SELECT NOME FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
      SQL.Active = True
      CurrentQuery.FieldByName("NOMERESPONSAVEL").AsString = SQL.FieldByName("NOME").AsString
    Case 2
      SQL.Clear
      SQL.Add("SELECT NOME FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
      SQL.Active = True
      CurrentQuery.FieldByName("NOMERESPONSAVEL").AsString = SQL.FieldByName("NOME").AsString
    Case 3
      SQL.Clear
      SQL.Add("SELECT NOME FROM SFN_PESSOA WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PESSOA").AsInteger
      SQL.Active = True
      CurrentQuery.FieldByName("NOMERESPONSAVEL").AsString = SQL.FieldByName("NOME").AsString
  End Select
  Set SQL = Nothing
End Sub

Public Sub TABLE_NewRecord()
  Dim vANO As String
  Dim SEQUENCIA As Long
  vANO = Format(ServerDate, "yyyy")
  NewCounter("CA_ATEND", CDate(vANO), 1, SEQUENCIA)
  CurrentQuery.FieldByName("ANO").Value = ("01/01/" + vANO)
  CurrentQuery.FieldByName("NUMERO").Value = SEQUENCIA
End Sub


Public Sub TABRESPONSAVEL_OnChange()



  'ESTADOPAGAMENTO.Visible=True
  'MUNICIPIOPAGAMENTO.Visible=True
  'ENDCORRESPONDENCIA.Visible=True
  'ENDATENDIMENTO.Visible=True
  'QTDVAGASESTACIONAMENTO.Visible=True
  'ENDERECOCOMERCIAL.Visible=True
  'ENDERECORESIDENCIAL.Visible=True
  'ENDERECOCORRESPONDENCIA.Visible=True
  'Select Case CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger
  '  Case 1
  '      ESTADOPAGAMENTO.Visible=False
  '      MUNICIPIOPAGAMENTO.Visible=False
  '      ENDCORRESPONDENCIA.Visible=False
  '      ENDATENDIMENTO.Visible=False
  '      QTDVAGASESTACIONAMENTO.Visible=False
  '  Case 2
  '      ENDERECOCOMERCIAL.Visible=False
  '      ENDERECORESIDENCIAL.Visible=False
  '      ENDERECOCORRESPONDENCIA.Visible=False
  '  Case 3
  '      ESTADOPAGAMENTO.Visible=False
  '      MUNICIPIOPAGAMENTO.Visible=False
  '      ENDCORRESPONDENCIA.Visible=False
  '      ENDATENDIMENTO.Visible=False
  '      QTDVAGASESTACIONAMENTO.Visible=False
  '      ENDERECOCOMERCIAL.Visible=False
  '      ENDERECORESIDENCIAL.Visible=False
  '      ENDERECOCORRESPONDENCIA.Visible=False
  'End Select
End Sub

'#######################################################
