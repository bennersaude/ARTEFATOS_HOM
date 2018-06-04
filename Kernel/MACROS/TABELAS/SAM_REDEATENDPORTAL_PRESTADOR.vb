'HASH: 66C8391BC950742B118B46F317B068B7

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim QPRESTADORDUP As Object
Set QPRESTADORDUP = NewQuery

QPRESTADORDUP.Active = False
QPRESTADORDUP.Clear
QPRESTADORDUP.Add("SELECT HANDLE                        ")
QPRESTADORDUP.Add("  FROM SAM_REDEATENDPORTAL_PRESTADOR ")
QPRESTADORDUP.Add(" WHERE REDEATENDPORTAL = :REDEATEND  ")
QPRESTADORDUP.Add("   AND PRESTADOR = :PRESTADOR        ")
QPRESTADORDUP.Add("   AND HANDLE <> :REGISTRO           ")
QPRESTADORDUP.ParamByName("REDEATEND").AsInteger = CurrentQuery.FieldByName("REDEATENDPORTAL").AsInteger
QPRESTADORDUP.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
QPRESTADORDUP.ParamByName("REGISTRO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
QPRESTADORDUP.Active = True

If Not QPRESTADORDUP.FieldByName("HANDLE").IsNull Then
  Set QPRESTADORDUP = Nothing
  CanContinue = False
  bsShowMessage("Prestador já cadastrado para esta rede de atendimento do Portal.","E")
  Exit Sub
End If

Set QPRESTADORDUP = Nothing

Dim msg As String
Dim NomeTabela As String
Dim CampoData1 As String
Dim CampoData2 As String
Dim DATAINICIAL As Date
Dim DATAFINAL As Date
Dim Campo As String
Dim Condicao As String

NomeTabela   = "SAM_REDEATENDPORTAL_PRESTADOR"
CampoData1   = "DataInicial"
CampoData2   ="DataFinal"
DATAINICIAL = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
DATAFINAL = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
Campo = "REDEATENDPORTAL"
Condicao = "AND PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString

Dim Interface As Object
Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
msg = Interface.Vigencia(CurrentSystem, NomeTabela, CampoData1, CampoData2, DATAINICIAL, DATAFINAL, Campo, Condicao, 0)

If msg <> "" Then
  CanContinue = False
  bsShowMessage(msg, "E")
End If

If ((Not CurrentQuery.FieldByName("DataFinal").IsNull) And (CurrentQuery.FieldByName("DataInicial").IsNull)) Then
    CanContinue = False
    bsShowMessage("Data Inicial obrigatória quando a Data Final é informada.","E")
    Exit Sub
End If


End Sub
