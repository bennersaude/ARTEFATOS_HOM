'HASH: 3086389CBF8E9A69EEE13F894BE639C7
 
'#Uses "*bsShowMessage"
Option Explicit
Public Sub CASAREPOUSO_OnChange()

	CurrentQuery.FieldByName("ENDERECO").Clear
	CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsBoolean = False
	CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsBoolean = False

End Sub

Public Sub ENDERECO_OnChange()

	Dim qTipoEnd As BPesquisa

	Set qTipoEnd = NewQuery

	qTipoEnd.Active = False
	qTipoEnd.Clear
	qTipoEnd.Add("SELECT CASE WHEN B.ENDERECOCPFCNPJ = :HENDERECO  ")
	qTipoEnd.Add("          THEN 'S' ELSE 'N' END ENDERECOCPFCNPJ, ")
	qTipoEnd.Add("       CASE WHEN B.ENDERECOCORRESPONDENCIA = :HENDERECO  ")
	qTipoEnd.Add("          THEN 'S' ELSE '' END ENDERECOCORRESPONDENCIA ")
	qTipoEnd.Add("  FROM SFN_PESSOA B ")
	qTipoEnd.Add(" WHERE B.Handle = :CASAREPOUSO ")
	qTipoEnd.ParamByName("HENDERECO").Value   = CurrentQuery.FieldByName("ENDERECO").AsInteger
	qTipoEnd.ParamByName("CASAREPOUSO").Value = CurrentQuery.FieldByName("CASAREPOUSO").AsInteger

	qTipoEnd.Active = True


	If (qTipoEnd.FieldByName("ENDERECOCPFCNPJ").AsString = "S") Then

		CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsBoolean = True

	Else

		CurrentQuery.FieldByName("ENDERECOCPFCNPJ").AsBoolean = False

	End If

	If (qTipoEnd.FieldByName("ENDERECOCORRESPONDENCIA").AsString = "S") Then

		CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsBoolean = True

	Else

		CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsBoolean = False

	End If

	qTipoEnd.Active = False
	Set qTipoEnd = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim NomeTabela As String
  Dim CampoData1 As String
  Dim CampoData2 As String
  Dim DataInicial As Date
  Dim DataFinal As Date
  Dim Condicao As String
  Dim HandleTabela As Integer
  Dim vsCritica As String

  NomeTabela   = "SAM_BENEFICIARIO_CASAREPOUSO"
  CampoData1   = "DATAENTRADA"
  CampoData2   = "DATASAIDA"
  DataInicial  = CurrentQuery.FieldByName("DATAENTRADA").AsDateTime
  DataFinal    = CurrentQuery.FieldByName("DATASAIDA").AsDateTime
  Condicao     = "BENEFICIARIO = "&CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  HandleTabela = CurrentQuery.FieldByName("HANDLE").AsInteger

  Dim dllVigencia As Object
  Set dllVigencia = CreateBennerObject("SAMGERAL.Vigencia")
  vsCritica = dllVigencia.Vigencia(CurrentSystem, NomeTabela, CampoData1, CampoData2, DataInicial, DataFinal, "", Condicao, HandleTabela)

  If Len(vsCritica) > 0 Then
    bsShowMessage(vsCritica, "E")
    CanContinue = False
  End If

End Sub
