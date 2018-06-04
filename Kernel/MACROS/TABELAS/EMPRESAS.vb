'HASH: 2B1EC12010DC38E233FA85913D496873
'Macro: EMPRESAS
'#Uses "*bsShowMessage"
Dim vNumeroContratoAutomaticoAnterior As String
Dim vNumeroFamiliaAutomaticoAnterior As String
Dim vNumeroBenefAutomaticoAnterior As String

Public Sub CEP_OnPopup(ShowPopup As Boolean)
  ' Joldemar Moreira 12/06/2003
  ' SMS 16059
  Dim vHandle As String
  Dim interface As Object
  ShowPopup = False
  Set interface = CreateBennerObject("ProcuraCEP.Rotinas")
  interface.Exec(CurrentSystem, vHandle)

  If vHandle <>"" Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO,COMPLEMENTO   ")
    SQL.Add("  FROM LOGRADOUROS      ")
    SQL.Add(" WHERE CEP = :HANDLE ")
    SQL.ParamByName("HANDLE").Value = vHandle
    SQL.Active = True

    CurrentQuery.Edit
    CurrentQuery.FieldByName("CEP").Value = SQL.FieldByName("CEP").AsString
    CurrentQuery.FieldByName("ESTADO").Value = SQL.FieldByName("ESTADO").AsString
    CurrentQuery.FieldByName("MUNICIPIO").Value = SQL.FieldByName("MUNICIPIO").AsString
    CurrentQuery.FieldByName("BAIRRO").Value = SQL.FieldByName("BAIRRO").AsString
    CurrentQuery.FieldByName("LOGRADOURO").Value = SQL.FieldByName("LOGRADOURO").AsString
  End If

  Set interface = Nothing

End Sub

Public Sub TABLE_AfterPost()
  vNumeroContratoAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROCONTRATOAUTOMATICO").AsString
  vNumeroFamiliaAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString
  vNumeroBenefAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString

  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "EMPRESAS")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_AfterScroll()

  vNumeroContratoAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROCONTRATOAUTOMATICO").AsString
  vNumeroFamiliaAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString
  vNumeroBenefAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString

  Dim vContratoNumAuto As String
  Dim vFamiliaNumAuto As String
  Dim vBenefNumAuto As String

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT * FROM SAM_PARAMETROSBENEFICIARIO")
  SQL.Active = True

  NUMEROCONTRATOAUTOMATICO.Visible = True
  NUMEROFAMILIAAUTOMATICO.Visible = True
  NUMEROBENEFAUTOMATICO.Visible = True

  Set SQL = Nothing

End Sub


Public Function CheckNumeroAutomatico As Boolean
  'verifica a troca do tipo de códigos
  Dim SQL1 As Object
  Set SQL1 = NewQuery
  Dim SQL2 As Object
  Set SQL2 = NewQuery
  Dim SQL3 As Object
  Set SQL3 = NewQuery

  SQL1.Add("SELECT COUNT(HANDLE) QTD FROM SAM_CONTRATO")
  SQL1.Active = True
  SQL2.Add("SELECT COUNT(HANDLE) QTD FROM SAM_FAMILIA")
  SQL2.Active = True
  SQL3.Add("SELECT COUNT(HANDLE) QTD FROM SAM_BENEFICIARIO")
  SQL3.Active = True


  'verifica a troca do tipo de códigos
  If NUMEROCONTRATOAUTOMATICO.Visible Then
    If(CurrentQuery.FieldByName("NUMEROCONTRATOAUTOMATICO").AsString = "S")And(vNumeroContratoAutomaticoAnterior = "N")Then
    If SQL1.FieldByName("QTD").AsInteger >0 Then
      bsShowMessage("Existem contratos cadastrados com número informado." + Chr(13) + "Impossível alterar o número do contrato para automático.", "E")
      CheckAlteracaoCodigos = False
      Set SQL1 = Nothing
      Set SQL2 = Nothing
      Set SQL3 = Nothing
      Exit Function
    End If
  End If
End If
If NUMEROFAMILIAAUTOMATICO.Visible Then
  If(CurrentQuery.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString = "S")And(vNumeroFamiliaAutomaticoAnterior = "N")Then
  If SQL2.FieldByName("QTD").AsInteger >0 Then
    bsShowMessage("Existem famílias cadastrados com número informado." + Chr(13) + "Impossível alterar o número da família para automático.", "E")
    CheckAlteracaoCodigos = False
    Set SQL1 = Nothing
    Set SQL2 = Nothing
    Set SQL3 = Nothing
    Exit Function
  End If
End If
End If
If NUMEROBENEFAUTOMATICO.Visible Then
  If(CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString = "S")And(vNumeroBenefAutomaticoAnterior = "N")Then
  If SQL3.FieldByName("QTD").AsInteger >0 Then
    bsShowMessage("Existem beneficiários cadastrados com número informado." + Chr(13) + "Impossível alterar o número do beneficiário para automático.", "E")
    CheckAlteracaoCodigos = False
    Set SQL1 = Nothing
    Set SQL2 = Nothing
    Set SQL3 = Nothing
    Exit Function
  End If
End If
End If

Set SQL1 = Nothing
Set SQL2 = Nothing
Set SQL3 = Nothing

CheckNumeroAutomatico = True

End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not IsValidCGC(CurrentQuery.FieldByName("CNPJ").AsString)Then
    bsShowMessage("CNPJ Inválido", E)
    CanContinue = False
    Exit Sub
  End If

  CanContinue = CheckNumeroAutomatico
  If CanContinue = False Then Exit Sub

  Dim vDep As Long
  Dim vFam As Long
  Dim vCont As Long
  Dim vEmp As Long
  Dim Interface As Object
  Set Interface = NewQuery
  Set Interface = CreateBennerObject("SamBeneficiario.Cadastro")
  Interface.ContaDigitosComposicaoBenef(CurrentSystem, vDep, vFam, vCont, vEmp)
  If(vEmp >0)And(Len(CurrentQuery.FieldByName("CODIGO").AsString)>vEmp)Then
  CanContinue = False
  bsShowMessage("Código da empresa com mais dígitos do que definido na composição do código do beneficiário.", "E")
  Set Interface = Nothing
  Exit Sub
End If

'#Uses "*VerificaEmail"

If Not CurrentQuery.FieldByName("EMAIL").IsNull Then
  If Not VerificaEmail(CurrentQuery.FieldByName("EMAIL").AsString)Then
    bsShowMessage("Endereço eletrônico inválido", "E")
    CanContinue = False
    Exit Sub
  End If
End If

Set Interface = Nothing
End Sub
