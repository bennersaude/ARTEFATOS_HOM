'HASH: 9919E7390736DD7B135883E445D33DF2
'#Uses "*bsShowMessage"

Public Function CHECARREDUCAO()
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim SQL As Object

  CHECARREDUCAO = True

  Condicao = " AND DATAFINAL >=  DATAINICIAL "

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_REDUCAOCAR", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATOMOD", Condicao)

  If Linha = "" Then
    CHECARREDUCAO = False
  Else
    CHECARREDUCAO = True
    bsShowMessage(Linha, "I")
  End If

  Set Interface = Nothing

End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CHECARREDUCAO Then
    CanContinue = False
    RefreshNodesWithTable("SAM_CONTRATO_REDUCAOCAR")
    Exit Sub
  End If


  'Daniela -13/11/2002
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Clear
  q1.Add("SELECT DATAADESAO FROM sAM_CONTRATO_MOD WHERE HANDLE = :HCONTRATOMOD")
  q1.ParamByName("HCONTRATOMOD").AsInteger = RecordHandleOfTable("SAM_CONTRATO_MOD")
  q1.Active = True
  If Not q1.EOF Then
    If(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <q1.FieldByName("DATAADESAO").AsDateTime)Then
    bsShowMessage("A data inicial da vigência não pode ser menor que a data de adesão do módulo", "E")
    CanContinue = False
  End If
End If
Set q1 = Nothing




End Sub


