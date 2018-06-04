'HASH: D0E66E1607509797827F7C23CB17D1AE
'Macro: SAM_BENEFICIARIO_REPASSE
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABLE_AfterPost()

  Dim UPD As Object
  Dim SQL As Object

  Set UPD = NewQuery
  Set SQL = NewQuery

  SQL.Add("SELECT * FROM SAM_BENEFICIARIO_REPASSE A WHERE A.BENEFICIARIO = :BENEF")
  SQL.Add("AND (DATAFINAL IS NULL OR DATAFINAL >= :DATAATUAL) ORDER BY HANDLE DESC")
  UPD.Add("UPDATE SAM_BENEFICIARIO SET CODIGODEREPASSE = :CODREPASSE WHERE HANDLE = :HANDLE")

  SQL.ParamByName("BENEF").Value = RecordHandleOfTable("SAM_BENEFICIARIO")
  SQL.ParamByName("DATAATUAL").Value = ServerDate
  SQL.Active = True

  If Not SQL.EOF Then

    UPD.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_BENEFICIARIO")
    UPD.ParamByName("CODREPASSE").Value = CurrentQuery.FieldByName("CODIGODEREPASSE").Value
    UPD.ExecSQL

  End If

  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = "AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_REPASSE", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", Condicao)

  If Linha <>"" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
  If CanContinue Then
    Check_Repasse CanContinue
  End If
  'verificar data final
  If CanContinue = True Then
    If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
      If CurrentQuery.FieldByName("DATAFINAL").AsDateTime <= ServerDate Then
        'limpar o código do repasse do beneficiário
        Dim SQL As Object
        Set SQL = NewQuery
        SQL.Add("SELECT CODIGODEREPASSE FROM SAM_BENEFICIARIO WHERE HANDLE = :HBENEF AND CODIGODEREPASSE = :CODREPASS")
        SQL.ParamByName("HBENEF").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
        SQL.ParamByName("CODREPASS").Value = CurrentQuery.FieldByName("CODIGODEREPASSE").AsInteger
        SQL.RequestLive = True
        SQL.Active = True
        If Not SQL.EOF Then
          SQL.Edit
          SQL.FieldByName("CODIGODEREPASSE").Clear
          SQL.Post
        End If
        Set SQL = Nothing
      End If
    End If
  End If
  If CanContinue = True Then
    CanContinue = CheckVigenciaBenef
  End If
End Sub

Public Sub Check_Repasse(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT COUNT(*) TREPAS FROM SAM_BENEFICIARIO_REPASSE A  WHERE A.CODIGODEREPASSE = :CODREPAS AND A.HANDLE <> :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("CODREPAS").Value = CurrentQuery.FieldByName("CODIGODEREPASSE").AsString
  SQL.Active = True
  If SQL.FieldByName("TREPAS").AsInteger >0 Then
    CanContinue = False
    bsShowMessage("Registro duplicado para CODIGO DE REPASSE do Beneficiário", "E")
  End If
  Set SQL = Nothing
End Sub

Public Function CheckVigenciaBenef As Boolean
  CheckVigenciaBenef = True
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DATAADESAO, DATACANCELAMENTO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEF")
  SQL.ParamByName("BENEF").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Data Inicial do Repasse inferior a Adesão do Beneficiário!", "I")
    CheckVigenciaBenef = False
  Else
    If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >SQL.FieldByName("DATACANCELAMENTO").AsDateTime Then
        bsShowMessage("Data de Inicial do Repasse maior que o cancelamento do Beneficiário !", "I")
        CheckVigenciaBenef = False
      End If
    End If
  End If
  Set SQL = Nothing
End Function

