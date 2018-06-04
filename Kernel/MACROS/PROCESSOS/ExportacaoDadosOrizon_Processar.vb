'HASH: B97FB8181B8A52A2C97BE9158219DBB1
'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"

Public Sub Main
  Dim qTisParametros       As Object
  Dim vDLLExportacaoOrizon As Object
  Dim qCodigoEMS           As Object
  Dim viHandleRotina       As Long
  Dim viChave              As Long
  Dim vsTabelasExportacao  As String

  If SessionVar("TABELASEXPORTACAOORIZON") = "" Then
    BsShowMessage("As tabelas para exportação não foram parametrizadas.", "E")
    Exit Sub
  End If

  Set qTisParametros = NewQuery

  qTisParametros.Clear
  qTisParametros.Add("SELECT DIRETORIOEXPORTACAOORIZON, DIRETORIOACESSOORIZON FROM TIS_PARAMETROS")
  qTisParametros.Active = True

  If qTisParametros.FieldByName("DIRETORIOEXPORTACAOORIZON").AsString = "" Then
    Set qTisParametros = Nothing
    BsShowMessage("O diretório para exportação não foi parametrizado.", "E")
    Exit Sub
  End If

  If qTisParametros.FieldByName("DIRETORIOACESSOORIZON").AsString = "" Then
    Set qTisParametros = Nothing
    BsShowMessage("O diretório para acesso não foi parametrizado. No caso do servidor do banco de dados ser hospedado em um ambiente Windows, copiar parametrização do campo Diretório Exportação. No caso do ambiente ser Linux, parametrizar com um mapeamento Windows do diretório do servidor Linux.", "E")
    Exit Sub
  End If

  CriaTabelaTemporariaSqlServer

  Set qCodigoEMS = NewQuery
  Set vDLLExportacaoOrizon = CreateBennerObject("Benner.Saude.Orizon.Exportacao.ExportacaoOrizon")

  qCodigoEMS.Clear
  qCodigoEMS.Add("SELECT CODIGOEMS FROM EMPRESAS WHERE HANDLE = :HANDLE")
  qCodigoEMS.ParamByName("HANDLE").AsInteger = CurrentCompany
  qCodigoEMS.Active = True

  NewCounter2("EXP_ORI_ROTINA", 0, 1, viChave)

  vsTabelasExportacao = SessionVar("TABELASEXPORTACAOORIZON")

  viHandleRotina = vDLLExportacaoOrizon.InserirRegistroRotinaExportacaoOrizon(CurrentSystem,viChave, vsTabelasExportacao, qTisParametros.FieldByName("DIRETORIOEXPORTACAOORIZON").AsString, qTisParametros.FieldByName("DIRETORIOACESSOORIZON").AsString, CurrentUser)

  vDLLExportacaoOrizon.ExportarDadosOrizon(CurrentSystem, viHandleRotina, viChave, qCodigoEMS.FieldByName("CODIGOEMS").AsString, SessionVar("TABELASEXPORTACAOORIZON"), qTisParametros.FieldByName("DIRETORIOEXPORTACAOORIZON").AsString, qTisParametros.FieldByName("DIRETORIOACESSOORIZON").AsString)

  Set qTisParametros = Nothing
  Set qCodigoEMS = Nothing
  Set vDLLExportacaoOrizon = Nothing
End Sub
