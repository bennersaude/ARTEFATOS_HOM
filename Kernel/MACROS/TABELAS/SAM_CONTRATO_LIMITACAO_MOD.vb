'HASH: 12D96EAA7D7DBA1086E0D82B75DDC9AB
'#Uses "*bsShowMessage"

Public Sub CONTRATOMODULO_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_CONTRATO_MOD|SAM_MODULO[SAM_CONTRATO_MOD.MODULO = SAM_MODULO.HANDLE]", "DESCRICAO", 1, "Descrição", "CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO")), "Procura por Módulo", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOMODULO").Value = handlexx
  End If
  Set Procura = Nothing

End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		CONTRATOMODULO.WebLocalWhere = "A.CONTRATO = " + Str(RecordHandleOfTable("SAM_CONTRATO"))
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim SQL As Object
  Dim SQL2 As Object

  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("Select LIMITACAO, CONTRATO, TIPOCONTAGEM, DATAINICIAL, DATAFINAL ")
  SQL.Add("  FROM SAM_CONTRATO_LIMITACAO  ")
  SQL.Add(" WHERE HANDLE = :HANDLE ")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_CONTRATO_LIMITACAO")
  SQL.Active = True

  Condicao = " AND LIMITACAO = " + SQL.FieldByName("LIMITACAO").AsString
  Condicao = Condicao + " AND TIPOCONTAGEM = '" + SQL.FieldByName("TIPOCONTAGEM").AsString + "'"

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_LIMITACAO", "DATAINICIAL", "DATAFINAL", SQL.FieldByName("DATAINICIAL").AsDateTime, SQL.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha = "" Then
    If(SQL.FieldByName("TIPOCONTAGEM").Value = "B")And _
       (SQL.FieldByName("INTERCAMBIAVEL").Value = "S")Then
    bsShowMessage("Acionar intercambiável somente para tipo de contagem contrato ou família", "E")
    CanContinue = False
  Else
    CanContinue = True
  End If
Else
  'Caso tenha outra limitação com vigência em aberto,permitir continuar caso não tenha limite sem módulo.

  Set SQL2 = NewQuery
    SQL2.Add("SELECT A.LIMITACAO, ")
    SQL2.Add("       B.CONTRATOMODULO ")
    SQL2.Add("  FROM SAM_CONTRATO_LIMITACAO A ")
    SQL2.Add("  Left Join SAM_CONTRATO_LIMITACAO_MOD B On B.CONTRATOLIMITACAO = A.HANDLE ")
    SQL2.Add(" WHERE A.CONTRATO = :CONTRATO ")
    SQL2.Add("   AND A.LIMITACAO = :LIMITACAO ")
    SQL2.Add("   AND B.CONTRATOMODULO = :MODULO ")
    SQL2.ParamByName("CONTRATO").Value = SQL.FieldByName("CONTRATO").AsInteger
    SQL2.ParamByName("LIMITACAO").Value = SQL.FieldByName("LIMITACAO").AsInteger
    SQL2.ParamByName("MODULO").Value = CurrentQuery.FieldByName("CONTRATOMODULO").AsInteger
  SQL2.Active = True

  If SQL2.EOF Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  End If

End If
'CanContinue =CheckVigencia
End Sub

