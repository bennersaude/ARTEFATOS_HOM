'HASH: E36282822235F36E81DEB0F537CA0262
'Macro: SAM_PARAMETROSBENEFICIARIO
'Juliano -03/01/2001
'Fábio -20/01/2002(Códigos)

'#Uses "*bsShowMessage"

Dim vNumeroContratoUnicoAnterior As Long
Dim vNumeroFamiliaUnicoAnterior As Long
Dim vNumeroBenefUnicoAnterior As Long
Dim vNumeroContratoAutomaticoAnterior As String
Dim vNumeroFamiliaAutomaticoAnterior As String
Dim vNumeroBenefAutomaticoAnterior As String
Dim gComposicaoCartao As String 'sms 28682
Dim gPermaneceNumeroCartao As Boolean 'sms 28682

Public Sub NUMEROBENEFUNICO_OnChange()
  NUMEROCONTRATOUNICO_OnChange
End Sub

Public Sub NUMEROCONTRATOUNICO_OnChange()
  If CurrentQuery.State <>1 Then CurrentQuery.UpdateRecord
  MostraEscondeAutomatico
End Sub

Public Sub NUMEROFAMILIAUNICO_OnChange()
  NUMEROCONTRATOUNICO_OnChange
End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  'sms 28682 início
  gComposicaoCartao = CurrentQuery.FieldByName("COMPOSICAOCARTAO").AsString
  gPermaneceNumeroCartao = CurrentQuery.FieldByName("PERMANECENUMEROCARTAO").AsString = "S"
  'sms 28682 fim

  Dim vDep As Long
  Dim vFam As Long
  Dim vCont As Long
  Dim vEmp As Long

  If CurrentQuery.FieldByName("ASSUMIRMATRICULA").AsString = "S" Then
    GRPHOMONIMO.Visible = True
  Else
    GRPHOMONIMO.Visible = False
  End If

  MostraEscondeAutomatico

  If(Not CurrentQuery.FieldByName("NUMEROCONTRATOUNICO").IsNull)And _
     (Not CurrentQuery.FieldByName("NUMEROFAMILIAUNICO").IsNull)And _
     (Not CurrentQuery.FieldByName("NUMEROBENEFUNICO").IsNull)Then

  vNumeroContratoUnicoAnterior = CurrentQuery.FieldByName("NUMEROCONTRATOUNICO").AsInteger
  vNumeroFamiliaUnicoAnterior = CurrentQuery.FieldByName("NUMEROFAMILIAUNICO").AsInteger
  vNumeroBenefUnicoAnterior = CurrentQuery.FieldByName("NUMEROBENEFUNICO").AsInteger

End If

Dim Interface As Object
Set Interface = NewQuery
Set Interface = CreateBennerObject("SamBeneficiario.Cadastro")
Interface.ContaDigitosComposicaoBenef(CurrentSystem, vDep, vFam, vCont, vEmp)
NUMERODEPENDENTEMAXIMO.Text = "Qtde máxima de dígitos para: emp." + Str(vEmp) + "; cont." + Str(vCont) + "; fam." + Str(vFam) + "; dep." + Str(vDep)
Set Interface = Nothing

End Sub

Public Sub MostraEscondeAutomatico

  If CurrentQuery.FieldByName("NUMEROBENEFUNICO").AsString = "5" Then
    GRPNUMERODEPENDENTE.Visible = True
  Else
    GRPNUMERODEPENDENTE.Visible = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'SMS 28682 INÍCIO
  Dim QAUX As Object
  Set QAUX = NewQuery
  QAUX.Clear
  If (InStr(SQLServer, "SQL") > 0) Or (InStr(SQLServer, "CACHE") > 0) Then
    QAUX.Add("SELECT TOP 1 HANDLE FROM SAM_BENEFICIARIO_CARTAOIDENTIF")
  ElseIf (InStr(SQLServer, "ORA") > 0) Then
    QAUX.Add("SELECT HANDLE FROM SAM_BENEFICIARIO_CARTAOIDENTIF WHERE ROWNUM = 1")
  Else
    QAUX.Add("SELECT MAX(HANDLE) HANDLE FROM SAM_BENEFICIARIO_CARTAOIDENTIF")
  End If
  QAUX.Active = True
  If (QAUX.FieldByName("HANDLE").AsInteger > 0) Then
    If (gComposicaoCartao <> CurrentQuery.FieldByName("COMPOSICAOCARTAO").AsString) Then
      bsShowMessage("Não é permitido alterar o parâmetro 'Composição do cartão' pois já existem cartões gerados no sistema.", "E")
      CanContinue = False
      Exit Sub
    End If
    If (gPermaneceNumeroCartao <> CurrentQuery.FieldByName("PERMANECENUMEROCARTAO").AsBoolean) Then
      bsShowMessage("Não é permitido alterar o parâmetro 'Permanece número cartão' pois já existem cartões gerados no sistema.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  Set QAUX = Nothing
  'SMS 28682 FIM

  Dim VNUMERODEPENDENTEINICIAL As Long

  VNUMERODEPENDENTEINICIAL = CurrentQuery.FieldByName("NUMERODEPENDENTEINICIAL").AsInteger
  If VNUMERODEPENDENTEINICIAL <>0 And VNUMERODEPENDENTEINICIAL <>1 Then
    bsShowMessage("O número do Dependente Inicial não pode ser diferente de 0 ou 1!", "E")
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABLEIAUTE").AsInteger = 1 Then

    'If(CurrentQuery.FieldByName("REGISTRO1").IsNull Or _
    '    CurrentQuery.FieldByName("REGISTRO2").IsNull Or _
    '    CurrentQuery.FieldByName("REGISTRO3").IsNull)Then
    '  CanContinue =False
    '  MsgBox ""

    'End If

    'If((CurrentQuery.FieldByName("REGISTRO1").AsString =CurrentQuery.FieldByName("REGISTRO2").AsString)Or _
    '     (CurrentQuery.FieldByName("REGISTRO1").AsString =CurrentQuery.FieldByName("REGISTRO3").AsString)Or _
    '     (CurrentQuery.FieldByName("REGISTRO2").AsString =CurrentQuery.FieldByName("REGISTRO3").AsString))Then
    '    CanContinue =False
    '    MsgBox "Indicadores idênticos."
    '    Exit Sub
    '  End If

  End If


  If CurrentQuery.FieldByName("TABLEIAUTE").AsInteger = 2 Then
    If CurrentQuery.FieldByName("SEPARADOR").IsNull Then
      CanContinue = False
      bsShowMessage("Separador obrigatório.", "E")
      Exit Sub
    End If

    If((CurrentQuery.FieldByName("IDENTIFCOBERTURA").AsString = CurrentQuery.FieldByName("IDENTIFCARENCIA").AsString)And _
       (Len(CurrentQuery.FieldByName("IDENTIFCOBERTURA").AsString)>0))Then
    CanContinue = False
    bsShowMessage("Idenficadores idênticos e não nulos.", "E")
    Exit Sub
  End If
End If

'SMS: 28682 [início]
If (Not CheckComposicaoBenef) Then
  CanContinue = False
  Exit Sub
End If

If (Not CheckComposicaoCartao) Then
  CanContinue = False
  Exit Sub
End If

CanContinue = CheckAlteracaoCodigos
If (CanContinue = False) Then
  Exit Sub
End If

If (ContaX(CurrentQuery.FieldByName("MASCARABENEFICIARIO").AsString) <> Len(CurrentQuery.FieldByName("COMPOSICAOBENEFICIARIO").AsString)) Then
  CanContinue = False
  bsShowMessage("Quantidade de 'X' na máscara não confere com a composição do beneficiário.", "E")
  Exit Sub
End If

'Compara a quantidade de X com o tamanho da composição do cartão acrescido de 1 ou de 2 por causa do dígito verificador.
If (Not((ContaX(CurrentQuery.FieldByName("MASCARABENEFICIARIOCARTAO").AsString) = Len(CurrentQuery.FieldByName("COMPOSICAOCARTAO").AsString) + 1) Or _
    (ContaX(CurrentQuery.FieldByName("MASCARABENEFICIARIOCARTAO").AsString) = Len(CurrentQuery.FieldByName("COMPOSICAOCARTAO").AsString) + 2))) Then
  CanContinue = False
  bsShowMessage("Quantidade de 'X' na máscara não confere com a composição do cartão." + Chr(13) + _
         "A máscara deve conter 1 ou 2 'X' a mais que a composição devido ao dígito verificador.", "E")
  Exit Sub
End If

If (CurrentQuery.FieldByName("CANCDEPOUTRAFAMILIA").AsString = "S") And (CurrentQuery.FieldByName("MOTIVOMIGRACAOCORRESPONSAVEL").IsNull) Then 'sms 60045
  bsShowMessage("Motivo para cancelamento do beneficiário - Migração de corresponsável deve ser informado", "E")
  CanContinue = False
  Exit Sub
End If

  If Not CurrentQuery.FieldByName("CLASSEGERENCIALFATPARC").IsNull Then
    Dim qClasseGerencial As Object
    Set qClasseGerencial = NewQuery
    qClasseGerencial.Active = False
    qClasseGerencial.Add("SELECT ULTIMONIVEL FROM SFN_CLASSEGERENCIAL WHERE HANDLE = :HCLASSEGERENCIAL")
    qClasseGerencial.ParamByName("HCLASSEGERENCIAL").AsInteger = CurrentQuery.FieldByName("CLASSEGERENCIALFATPARC").AsInteger
    qClasseGerencial.Active = True

    If qClasseGerencial.FieldByName("ULTIMONIVEL").AsString <> "S" Then
      bsShowMessage("A classe gerencial da fatura saldo parcelamento escolhida deve ser de último nível.", "E")
      CanContinue = False
      Set qClasseGerencial = Nothing
      Exit Sub
    End If
    Set qClasseGerencial = Nothing
  End If

If ((CurrentQuery.FieldByName("ASSUMIRMATRICULA").AsString = "N") And (CurrentQuery.FieldByName("ASSUMIRMATRICULAOUTROCONTRATO").AsString = "S"))Then
  bsShowMessage("O flag Assumir Matrícula de Outro Contrato só pode ser marcado se o flag Assumir  matrícula existente automático estiver marcado.", "E")
  CanContinue = False
  Exit Sub
End If

  Dim qBuscaRelatorioProvisorioCodigoDuplicado As Object

  Set qBuscaRelatorioProvisorioCodigoDuplicado = NewQuery
  qBuscaRelatorioProvisorioCodigoDuplicado.Active = False
  qBuscaRelatorioProvisorioCodigoDuplicado.Add("SELECT COUNT(HANDLE) QUANT FROM R_RELATORIOS WHERE CODIGO = :CODIGO")
  qBuscaRelatorioProvisorioCodigoDuplicado.ParamByName("CODIGO").AsString = CurrentQuery.FieldByName("RELATORIOIMPRESSAOPROVISORIA").AsString
  qBuscaRelatorioProvisorioCodigoDuplicado.Active = True

  If (qBuscaRelatorioProvisorioCodigoDuplicado.FieldByName("QUANT").AsInteger > 1) Then
	CanContinue = False
	BsShowMessage("Não é permitido Relatório para Impressão Provisória com código duplicado!", "E")
  End If

End Sub

Public Function ContaX(pStr As String)As Integer
  Dim i As Integer
  Dim cont As Integer

  For i = 1 To Len(pStr)
    If Mid(pStr, i, 1) = "X" Then cont = cont + 1
  Next i
  ContaX = cont

End Function

'SMS: 28682 [início]

Public Function CheckComposicaoBenef As Boolean
  Dim vOK As Boolean
  Dim Interface As Object

  Set Interface = CreateBennerObject("SamBeneficiario.Cadastro")

  Interface.ValidaMascara(CurrentSystem, _
                          CurrentQuery.FieldByName("COMPOSICAOBENEFICIARIO").AsString, _
                          "B", _
                          CurrentQuery.FieldByName("IDEMPRESA").AsString, _
                          CurrentQuery.FieldByName("IDCONTRATO").AsString, _
                          CurrentQuery.FieldByName("IDFAMILIA").AsString, _
                          CurrentQuery.FieldByName("IDDEPENDENTE").AsString, _
                          CurrentQuery.FieldByName("IDCARTAO").AsString, _
                          CurrentQuery.FieldByName("NUMEROCONTRATOUNICO").AsString, _
                          CurrentQuery.FieldByName("NUMEROFAMILIAUNICO").AsString, _
                          CurrentQuery.FieldByName("NUMEROBENEFUNICO").AsString, _
                          CurrentQuery.FieldByName("IDVIA").AsString, _
                          vOK)

  If (Not vOK) Then
    bsShowMessage("Composição do código do beneficiário inválida.", "E")
    CheckComposicaoBenef = False
  Else
    CheckComposicaoBenef = True
  End If

  Set Interface = Nothing
End Function

Public Function CheckComposicaoCartao As Boolean
  Dim vOK As Boolean
  Dim Interface As Object

  Set Interface = CreateBennerObject("SamBeneficiario.Cadastro")


  Interface.ValidaMascara(CurrentSystem, _
                          CurrentQuery.FieldByName("COMPOSICAOCARTAO").AsString, _
                          "C", _
                          CurrentQuery.FieldByName("IDEMPRESA").AsString, _
                          CurrentQuery.FieldByName("IDCONTRATO").AsString, _
                          CurrentQuery.FieldByName("IDFAMILIA").AsString, _
                          CurrentQuery.FieldByName("IDDEPENDENTE").AsString, _
                          CurrentQuery.FieldByName("IDCARTAO").AsString, _
                          CurrentQuery.FieldByName("NUMEROCONTRATOUNICO").AsString, _
                          CurrentQuery.FieldByName("NUMEROFAMILIAUNICO").AsString, _
                          CurrentQuery.FieldByName("NUMEROBENEFUNICO").AsString, _
                          CurrentQuery.FieldByName("IDVIA").AsString, _
                          vOK)

  If (Not vOK) Then
    bsShowMessage("Composição do código do cartão inválida.", "E")
    CheckComposicaoCartao = False
  Else
    CheckComposicaoCartao = True
  End If

  Set Interface = Nothing
End Function

'SMS: 28682 [fim]

Public Function CheckAlteracaoCodigos As Boolean

  CurrentQuery.UpdateRecord

  If(CurrentQuery.FieldByName("NUMEROCONTRATOUNICO").AsInteger = 3)Or _
     (CurrentQuery.FieldByName("NUMEROFAMILIAUNICO").AsInteger = 3)Or _
     (CurrentQuery.FieldByName("NUMEROBENEFUNICO").AsInteger = 3)Then
  bsShowMessage("Convênio ainda não implementado", "E")
  CheckAlteracaoCodigos = False
  Exit Function
End If


'verifica a troca do tipo de códigos
Dim SQL1 As Object
Set SQL1 = NewQuery
Dim SQL2 As Object
Set SQL2 = NewQuery
Dim SQL3 As Object
Set SQL3 = NewQuery

'somente pode p/baixo,p/cima não se existir algo cadastrado.
SQL1.Add("SELECT MAX(HANDLE) QTD FROM SAM_CONTRATO")
SQL1.Active = True
If(CurrentQuery.FieldByName("NUMEROCONTRATOUNICO").AsInteger <vNumeroContratoUnicoAnterior)And(SQL1.FieldByName("QTD").AsInteger >0)Then
  bsShowMessage("Existem contratos cadastrados. " + Chr(13) + "Impossível alterar o tipo do número do contrato.", "E")
  CheckAlteracaoCodigos = False
  Set SQL1 = Nothing
  Set SQL2 = Nothing
  Set SQL3 = Nothing
  Exit Function
End If
SQL2.Add("SELECT MAX(HANDLE) QTD FROM SAM_FAMILIA")
SQL2.Active = True
If(CurrentQuery.FieldByName("NUMEROFAMILIAUNICO").AsInteger <vNumeroFamiliaUnicoAnterior)And(SQL2.FieldByName("QTD").AsInteger >0)Then
  bsShowMessage("Existem famílias cadastradas. " + Chr(13) + "Impossível alterar o tipo do número da família.", "E")
  CheckAlteracaoCodigos = False
  Set SQL1 = Nothing
  Set SQL2 = Nothing
  Set SQL3 = Nothing
  Exit Function
End If
SQL3.Add("SELECT MAX(HANDLE) QTD FROM SAM_BENEFICIARIO")
SQL3.Active = True
If(CurrentQuery.FieldByName("NUMEROBENEFUNICO").AsInteger <vNumeroBenefUnicoAnterior)And(SQL3.FieldByName("QTD").AsInteger >0)Then
  bsShowMessage("Existem beneficiários cadastrados. " + Chr(13) + "Impossível alterar o tipo do código Do beneficiário.", "E")
  CheckAlteracaoCodigos = False
  Set SQL1 = Nothing
  Set SQL2 = Nothing
  Set SQL3 = Nothing
  Exit Function
End If



Set SQL1 = Nothing
Set SQL2 = Nothing
Set SQL3 = Nothing

CheckAlteracaoCodigos = True

End Function

Public Sub RELATORIOIMPRESSAOPROVISORIA_OnBtnClick()
  Dim OLEAutorizador As Object
  Dim handlexx As Long
  On Error GoTo cancel
  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
  handlexx = OLEAutorizador.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Relatório|Código", "", "Procura por Relatórios", True, "")
  If handlexx <>0 Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CODIGO FROM R_RELATORIOS WHERE HANDLE =           :HANDLE")
    SQL.ParamByName("HANDLE").Value = handlexx
    SQL.Active = True
    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("RELATORIOIMPRESSAOPROVISORIA").Value = SQL.FieldByName("CODIGO").AsString
  End If
  Set OLEAutorizador = Nothing
Cancel :
End Sub

