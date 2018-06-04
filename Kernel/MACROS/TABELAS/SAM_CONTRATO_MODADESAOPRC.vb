'HASH: 1ADF8BB826E943A801E0666B1C35CE6F
'Macro: SAM_CONTRATO_MODADESAOPRC
'#Uses "*bsShowMessage"


Dim vPercentualValorInscricaoAnterior As String
Dim vValorInscricao As Double
Dim vAlteracao As Boolean

Public Sub BOTAOIMPORTARFAIXAS_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If
  Dim INTERFACE0002 As Object
  Dim vcContainer As CSDContainer
  Set vcContainer = NewContainer

  SessionVar("TIPOIMPORTAFAIXA") = "C"

  Dim vsMensagem As String
  If VisibleMode Then
  	Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  	INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0105", _
					   "Importação de Tabela de Preço",  _
					   0, _
					   400, _
					   530, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  End If

  SessionVar("TIPOIMPORTAFAIXA") = ""
  Set INTERFACE0002 = Nothing
  Set vcContainer = Nothing
End Sub

Public Sub TABLE_AfterPost()
  Dim vPrimeiraCompetencia As Date
  Dim vUltimaCompetencia As Date
  Dim SQL As Object

  If vAlteracao And _
      (CurrentQuery.FieldByName("PERCENTUALVALORINSCRICAO").AsString <>vPercentualValorInscricaoAnterior Or _
      CurrentQuery.FieldByName("VALORINSCRICAO").AsString <>vValorInscricaoAnterior)Then

    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT MIN(COMPETENCIA) PRIMEIRACOMPETENCIA, MAX(COMPETENCIA) ULTIMACOMPETENCIA")
    SQL.Add("FROM SFN_FATURA_LANC_MOD A, SAM_BENEFICIARIO_MOD BM,")
    SQL.Add("     SAM_CONTRATO_MOD CM")
    SQL.Add("WHERE CM.HANDLE = :HCONTRATOMOD")
    SQL.Add("  AND BM.MODULO = CM.HANDLE")
    SQL.Add("  AND (   BM.DATACANCELAMENTO IS NULL")
    SQL.Add("       OR BM.DATACANCELAMENTO >= :HOJE)")
    SQL.Add("  AND A.BENEFICIARIOMOD = BM.HANDLE")
    SQL.Add("  AND A.TABTIPO = 3")
    SQL.Add("  AND A.COMPETENCIA >= :COMPETENCIA")
    SQL.ParamByName("HCONTRATOMOD").Value = RecordHandleOfTable("SAM_CONTRATO_MOD")
    SQL.ParamByName("HOJE").Value = ServerDate
    SQL.ParamByName("COMPETENCIA").Value = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
    SQL.Active = True

    vPrimeiraCompetencia = SQL.FieldByName("PRIMEIRACOMPETENCIA").AsDateTime
    vUltimaCompetencia = SQL.FieldByName("ULTIMACOMPETENCIA").AsDateTime

    If Not SQL.FieldByName("PRIMEIRACOMPETENCIA").IsNull Then

      SQL.Clear
      SQL.Add("INSERT INTO SAM_ROTINARECALCULOMENSALID")
      SQL.Add("(HANDLE, CODIGO, DESCRICAO, DATAROTINA, TABRECALCULAR,")
      SQL.Add(" COMPETENCIAINICIAL, COMPETENCIAFINAL, CONTRATOINICIAL, CONTRATOFINAL,")
      SQL.Add(" USUARIO, DATAINCLUSAO, SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO)")
      SQL.Add("VALUES")
      SQL.Add("(:HANDLE, :HANDLE, :DESCRICAO, :DATAROTINA, 2,")
      SQL.Add(" :COMPETENCIAINICIAL, :COMPETENCIAFINAL, :HCONTRATO, :HCONTRATO,")
      SQL.Add(" :USUARIO, :DATAINCLUSAO, '1', '1')")

      SQL.ParamByName("HANDLE").Value = NewHandle("SAM_ROTINARECALCULOMENSALID")
      SQL.ParamByName("DATAROTINA").Value = ServerDate
      SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
      SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
      SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
      SQL.ParamByName("USUARIO").Value = CurrentUser
      SQL.ParamByName("DATAINCLUSAO").Value = ServerDate
      SQL.ParamByName("DESCRICAO").Value = "Alteração na configuração da Taxa de Inscrição"

      SQL.ExecSQL

    End If
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vAlteracao = True
  vPercentualValorInscricaoAnterior = CurrentQuery.FieldByName("PERCENTUALVALORINSCRICAO").AsString
  vValorInscricao = CurrentQuery.FieldByName("VALORINSCRICAO").AsFloat
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  vAlteracao = False
  Dim SQL
  Set SQL = NewQuery
  SQL.Add("SELECT DATACANCELAMENTO FROM SAM_CONTRATO_MOD WHERE HANDLE = :HCONTRATOMOD")
  SQL.ParamByName("HCONTRATOMOD").Value = RecordHandleOfTable("SAM_CONTRATO_MOD")
  SQL.Active = True
  If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("Módulo cancelado não permite manutenções", "I")
    CurrentQuery.Cancel
    RefreshNodesWithTable("SAM_CONTRATO_MODADESAOPRC")
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qFechamento As Object
  Set qFechamento = NewQuery
  Dim vMesFechamento As Integer
  Dim vAnoFechamento As Integer
  Dim vAnoComp As Integer
  Dim vmesComp As Integer

  qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  qFechamento.Active = True

  vMesFechamento = DatePart("m", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
  vAnoFechamento = DatePart("yyyy", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

  vmesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
  vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)

  If CurrentQuery.State = 3 Then
    If(vAnoComp <vAnoFechamento)Or _
       (vAnoComp = vAnoFechamento And vmesComp <vMesFechamento)Then
    CanContinue = False
    bsShowMessage("A competência não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
  End If
End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOIMPORTARFAIXAS" Then
		BOTAOIMPORTARFAIXAS_OnClick
	End If
End Sub
