'HASH: 0C23F12A2B2F5443B3E36E4B05FBB603
'Macro: SAM_MIGRACAO_BENEF_LIMITACAO
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If

  'SMS 61198 - Matheus - Início
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SL.PERIODICIDADE          ")
  SQL.Add("  FROM SAM_LIMITACAO SL,         ")
  SQL.Add("       SAM_CONTRATO_LIMITACAO SCL")
  SQL.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  SQL.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODOCONTAGEM.Visible = False
  Else
    PERIODOCONTAGEM.Visible = True
  End If

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = " AND LIMITACAO = " + CurrentQuery.FieldByName("LIMITACAO").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_LIMITACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If
  CanContinue = CheckVigenciaBenef
End Sub

Public Function CheckVigenciaBenef As Boolean
  CheckVigenciaBenef = True
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DATAADESAO,DATACANCELAMENTO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEF")
  SQL.ParamByName("BENEF").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Data Inicial da Limitação inferior a Adesão do Beneficiário!", "E")
    CheckVigenciaBenef = False
  Else
    If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >SQL.FieldByName("DATACANCELAMENTO").AsDateTime Then
        bsShowMessage("Data de Inicial da Limitação maior que o cancelamento do Beneficiário !", "E")
        CheckVigenciaBenef = False
      End If
    End If
  End If
  Set SQL = Nothing
End Function

