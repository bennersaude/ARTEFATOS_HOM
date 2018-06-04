'HASH: FEC399AE29982002D4C5FD85015F5665

Public Sub Main
  Dim msg As String
  Dim NomeTabela As String
  Dim CampoData1 As String
  Dim CampoData2 As String
  Dim DataInicial As Date
  Dim DataFinal As Date
  Dim Condicao As String
  Dim HandleTabela As Long

  NomeTabela   = CStr(ServiceVar("NOMETABELA"))
  CampoData1   = CStr(ServiceVar("CAMPODATA1"))
  CampoData2   = CStr(ServiceVar("CAMPODATA2"))
  DataInicial  = CDate(ServiceVar("DATAINICIAL"))
  DataFinal    = CDate(ServiceVar("DATAFINAL"))
  Condicao     = CStr(ServiceVar("CONDICAO"))
  HandleTabela = CLng(ServiceVar("HANDLETABELA"))

  Dim Interface As Object
  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  msg = Interface.Vigencia(CurrentSystem, NomeTabela, CampoData1, CampoData2, DataInicial, DataFinal, "", Condicao, HandleTabela)

  Set Interface = Nothing

  ServiceResult = msg

End Sub
