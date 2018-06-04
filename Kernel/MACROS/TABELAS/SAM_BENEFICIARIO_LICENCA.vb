'HASH: 279AA4AD24C9C31B4990C4547A4F24A3
'Macro: SAM_BENEFICIARIO_LICENCA
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If

  If CurrentQuery.FieldByName("SITUACAOSOLICITPERMANENCIA").IsNull Then
    GRPINFOSOLICITPERMANENCIA.Visible = False
  Else
    GRPINFOSOLICITPERMANENCIA.Visible = True
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CanContinue = False
    bsSHowMessage("Registro finalizado não pode ser alterado!", "E")
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_LICENCA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", "")

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
  Set Interface = Nothing
  CanContinue = CheckVigenciaBenef

  If CurrentQuery.FieldByName("PARTICULAR").AsString = _
                              CurrentQuery.FieldByName("ACIDENTETRABALHO").AsString Then
    If CurrentQuery.FieldByName("PARTICULAR").AsString = "S" Then
      CanContinue = False
      bsShowMessage("'Particular' e 'Acidente de trabalho' não podem ser marcados ao mesmo tempo", "E")
    End If
  End If

  Dim SQLFechamento
  Set SQLFechamento = NewQuery
  SQLFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  SQLFechamento.Active = True

  If Not CurrentQuery.FieldByName("DATAINICIAL").IsNull Then

    If SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime >CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
      bsShowMessage("Não é possível cadastrar data inicial inferior a data de fechamento - Parâmetros Gerais", "E")
      CanContinue = False
    End If
  End If
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime >CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
      bsShowMessage("Não é possível cadastrar data final inferior a data de fechamento - Parâmetros Gerais", "E")
      CanContinue = False
    End If
  End If


  Set SQLFechamento = Nothing


End Sub

Public Function CheckVigenciaBenef As Boolean
  CheckVigenciaBenef = True
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DATAADESAO,DATACANCELAMENTO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEF")
  SQL.ParamByName("BENEF").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Data Inicial da licença inferior a Adesão do Beneficiário!", "E")
    CheckVigenciaBenef = False
  Else
    If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >SQL.FieldByName("DATACANCELAMENTO").AsDateTime Then
        bsShowMessage("Data de Inicial da licença maior que o cancelamento do Beneficiário !", "E")
        CheckVigenciaBenef = False
      End If
    End If
  End If
  Set SQL = Nothing
End Function

