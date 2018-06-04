'HASH: 7386E039D2B865CD75868168526A4730
'SAM_BENEFICIARIO_SALARIO'
Option Explicit

Public Function ProcuraBeneficiarioAtivo(pSoAtivos As Boolean, pData As Date, TextoBenAtivo As String) As Long
  Dim interface As Object
  Dim vWhere As String
  Dim vColunas As String
  Dim qparametros As BPesquisa
  Set qparametros = NewQuery


  qparametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
  qparametros.Active = True
  If qparametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString = "S" Then
    'Set interface = CreateBennerObject("CA010.ConsultaBeneficiario")
    'Alterado SMS 90338 - Rodrigo Andrade 30/11/2007 -
    'Separação da Interface da regra de negocio para consulta de Beneficiários
    Set interface =CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")
    ProcuraBeneficiarioAtivo = interface.Filtro(CurrentSystem, 1, "")
    Set interface = Nothing
  End If

  If qparametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString = "N" Then
    vColunas = "SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_CONTRATO.CONTRATANTE|SAM_BENEFICIARIO.CODIGODEAFINIDADE|SAM_BENEFICIARIO.CODIGOANTIGO|SAM_BENEFICIARIO.DATACANCELAMENTO|SAM_CONVENIO.DESCRICAO|SAM_BENEFICIARIO.CODIGODEORIGEM|SAM_BENEFICIARIO.CODIGODEREPASSE"

    vWhere = ""

    If pSoAtivos = True Then

      vWhere = vWhere + "(SAM_BENEFICIARIO.DATABLOQUEIO IS NULL) And "
      vWhere = vWhere + " ((SAM_BENEFICIARIO.ATENDIMENTOATE Is NOT NULL AND SAM_BENEFICIARIO.ATENDIMENTOATE >= " + SQLDate(pData) + ") OR (SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL OR SAM_BENEFICIARIO.DATACANCELAMENTO >= " + SQLDate(pData) + "))"
    End If

    Set interface = CreateBennerObject("Procura.Procurar")
    ProcuraBeneficiarioAtivo = interface.Exec("SAM_BENEFICIARIO|SAM_CONTRATO[SAM_BENEFICIARIO.CONTRATO=SAM_CONTRATO.HANDLE]|SAM_CONVENIO[SAM_BENEFICIARIO.CONVENIO=SAM_CONVENIO.HANDLE]", vColunas, 2, "Nome|Beneficiario|Contratante|Código Afinidade|Código Antigo|Data Cancelamento|Convenio|Código de origem|Código de repasse", vWhere, "Procura por Beneficiário", False, TextoBenAtivo , "CA006.ConsultaBeneficiario")
    Set interface = Nothing
  End If

  Set qparametros = Nothing
End Function

Public Sub TABLE_AfterCancel()
	If WebMode Then
		BENEFICIARIO.WebLocalWhere = ""

	ElseIf VisibleMode Then
		BENEFICIARIO.LocalWhere = ""
	End If
End Sub

Public Sub TABLE_AfterPost()
  Dim qSQL As BPesquisa
  Set qSQL = NewQuery

  Dim qContratoSalario As BPesquisa
  Set qContratoSalario = NewQuery

  qContratoSalario.Clear
  qContratoSalario.Add("SELECT SEQUENCIA FROM SAM_CONTRATO_SALARIO WHERE HANDLE = :HCONTRATOSALARIO")
  qContratoSalario.ParamByName("HCONTRATOSALARIO").Value = CurrentQuery.FieldByName("CONTRATOSALARIO").AsInteger
  qContratoSalario.Active = True

  If qContratoSalario.FieldByName("SEQUENCIA").AsString <> CurrentQuery.FieldByName("SEQUENCIA").AsString Then
     qSQL.Clear
     qSQL.Add("UPDATE SAM_BENEFICIARIO_SALARIO SET SEQUENCIA = :SEQUENCIA WHERE HANDLE =:HBENSALARIO")
     qSQL.ParamByName("SEQUENCIA").AsString    = CONTRATOSALARIO.Text
     qSQL.ParamByName("HBENSALARIO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
     qSQL.ExecSQL
  End If

  Set qContratoSalario = Nothing
  Set qSQL = Nothing

  RefreshNodesWithTable("SAM_BENEFICIARIO_SALARIO")

	If WebMode Then
		BENEFICIARIO.WebLocalWhere = ""

	ElseIf VisibleMode Then
		BENEFICIARIO.LocalWhere = ""
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		BENEFICIARIO.WebLocalWhere = "A.HANDLE IN (SELECT A.HANDLE                          " + _
                              "             FROM SAM_BENEFICIARIO                " + _
                              "            WHERE CONTRATO = (SELECT CONTRATO FROM SAM_CONTRATO_SALARIO WHERE HANDLE = "+  CurrentQuery.FieldByName("CONTRATOSALARIO").AsString + "))"

	ElseIf VisibleMode Then
	    BENEFICIARIO.LocalWhere = "A.HANDLE IN (SELECT HANDLE                          " + _
                              "             FROM SAM_BENEFICIARIO                " + _
                              "            WHERE CONTRATO = (SELECT CONTRATO FROM SAM_CONTRATO_SALARIO WHERE HANDLE = "+  CurrentQuery.FieldByName("CONTRATOSALARIO").AsString + "))"
    End If


    Dim qContrato As BPesquisa
  	Set qContrato = NewQuery

  	If CurrentQuery.FieldByName("BENEFICIARIO").AsInteger > 0 Then
	  	qContrato.Active = False
	  	qContrato.Clear
	  	qContrato.Add("SELECT B.CONTRATO HCONTRATO ")
	  	qContrato.Add("  FROM SAM_BENEFICIARIO B   ")
	  	qContrato.Add(" WHERE B.HANDLE = :HBENEF   ")
	  	qContrato.ParamByName("HBENEF").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
	  	qContrato.Active = True
	Else
	  	qContrato.Active = False
	  	qContrato.Clear
	  	qContrato.Add("SELECT CONTRATO HCONTRATO ")
	  	qContrato.Add("  FROM SAM_CONTRATO_SALARIO    ")
	  	qContrato.Add(" WHERE HANDLE = :HCONTRATOSALARIO   ")
	  	qContrato.ParamByName("HCONTRATOSALARIO").AsInteger = CurrentQuery.FieldByName("CONTRATOSALARIO").AsInteger
	  	qContrato.Active = True

	End If

    If WebMode Then
    	CONTRATOSALARIO.WebLocalWhere = " A.CONTRATO = " + qContrato.FieldByName("HCONTRATO").AsString + _
                               		" AND NOT EXISTS (SELECT 1                                                " + _
                               		"                   FROM SAM_BENEFICIARIO_SALARIO B                       " + _
                               		"                  WHERE B.BENEFICIARIO    = @CAMPO(BENEFICIARIO)  AND B.CONTRATOSALARIO = A.HANDLE) "
    ElseIf VisibleMode Then
  		CONTRATOSALARIO.LocalWhere = " CONTRATO = " + qContrato.FieldByName("HCONTRATO").AsString + _
                               		" AND NOT EXISTS (SELECT 1                                                " + _
                               		"                   FROM SAM_BENEFICIARIO_SALARIO B                       " + _
                               		"                  WHERE B.BENEFICIARIO    = @BENEFICIARIO AND B.CONTRATOSALARIO = SAM_CONTRATO_SALARIO.HANDLE) "
    End If
  	Set qContrato = Nothing



End Sub

Public Sub TABLE_AfterInsert()
	If WebMode Then
		BENEFICIARIO.WebLocalWhere = "A.HANDLE IN (SELECT A.HANDLE                          " + _
                              "             FROM SAM_BENEFICIARIO                " + _
                              "            WHERE CONTRATO = (SELECT CONTRATO FROM SAM_CONTRATO_SALARIO WHERE HANDLE = "+  CurrentQuery.FieldByName("CONTRATOSALARIO").AsString + "))"

	ElseIf VisibleMode Then
	    BENEFICIARIO.LocalWhere = "A.HANDLE IN (SELECT HANDLE                          " + _
                              "             FROM SAM_BENEFICIARIO                " + _
                              "            WHERE CONTRATO = (SELECT CONTRATO FROM SAM_CONTRATO_SALARIO WHERE HANDLE = "+  CurrentQuery.FieldByName("CONTRATOSALARIO").AsString + "))"
    End If


    Dim qContrato As BPesquisa
  	Set qContrato = NewQuery

  	If CurrentQuery.FieldByName("BENEFICIARIO").AsInteger > 0 Then
	  	qContrato.Active = False
	  	qContrato.Clear
	  	qContrato.Add("SELECT B.CONTRATO HCONTRATO ")
	  	qContrato.Add("  FROM SAM_BENEFICIARIO B   ")
	  	qContrato.Add(" WHERE B.HANDLE = :HBENEF   ")
	  	qContrato.ParamByName("HBENEF").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
	  	qContrato.Active = True
	Else
	  	qContrato.Active = False
	  	qContrato.Clear
	  	qContrato.Add("SELECT CONTRATO HCONTRATO ")
	  	qContrato.Add("  FROM SAM_CONTRATO_SALARIO    ")
	  	qContrato.Add(" WHERE HANDLE = :HCONTRATOSALARIO   ")
	  	qContrato.ParamByName("HCONTRATOSALARIO").AsInteger = CurrentQuery.FieldByName("CONTRATOSALARIO").AsInteger
	  	qContrato.Active = True

	End If

    If WebMode Then
    	CONTRATOSALARIO.WebLocalWhere = " A.CONTRATO = " + qContrato.FieldByName("HCONTRATO").AsString + _
                               		" AND NOT EXISTS (SELECT 1                                                " + _
                               		"                   FROM SAM_BENEFICIARIO_SALARIO B                       " + _
                               		"                  WHERE B.BENEFICIARIO    = @CAMPO(BENEFICIARIO)  AND B.CONTRATOSALARIO = A.HANDLE) "
    ElseIf VisibleMode Then
  		CONTRATOSALARIO.LocalWhere = " CONTRATO = " + qContrato.FieldByName("HCONTRATO").AsString + _
                               		" AND NOT EXISTS (SELECT 1                                                " + _
                               		"                   FROM SAM_BENEFICIARIO_SALARIO B                       " + _
                               		"                  WHERE B.BENEFICIARIO    = @BENEFICIARIO AND B.CONTRATOSALARIO = SAM_CONTRATO_SALARIO.HANDLE) "
    End If
  	Set qContrato = Nothing

End Sub

Public Sub TABLE_UpdateRequired()
  If (Not VisibleMode) And Not(CurrentQuery.FieldByName("CONTRATOSALARIO").IsNull) Then
    Dim qContratoSalario As BPesquisa
    Set qContratoSalario = NewQuery

    qContratoSalario.Add("SELECT SEQUENCIA FROM SAM_CONTRATO_SALARIO WHERE HANDLE = :HCONTRATOSALARIO")
    qContratoSalario.ParamByName("HCONTRATOSALARIO").Value = CurrentQuery.FieldByName("CONTRATOSALARIO").AsInteger
    qContratoSalario.Active = True

    CurrentQuery.FieldByName("SEQUENCIA").AsString = qContratoSalario.FieldByName("SEQUENCIA").AsString

    Set qContratoSalario = Nothing
  End If
End Sub
