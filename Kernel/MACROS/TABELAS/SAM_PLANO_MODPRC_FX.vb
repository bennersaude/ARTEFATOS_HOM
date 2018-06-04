'HASH: 9023A038DDF8A0FA113D6BE57D7B5DC2
'SAM_PLANO_MODPRC_FX
'#Uses "*bsShowMessage"

Dim InclusaoForaDoPadrao As Boolean
Dim LogInclusaoForaDoPadrao As String

Public Sub PLANOTPDEP_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_PLANO_TPDEP|SAM_TIPODEPENDENTE[SAM_PLANO_TPDEP.TIPODEPENDENTE = SAM_TIPODEPENDENTE.HANDLE]", "DESCRICAO", 1, "Descrição", "PLANO = " + Str(RecordHandleOfTable("SAM_PLANO")), "Procura por Tipo dependente", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PLANOTPDEP").Value = handlexx
  End If
  Set Procura = Nothing
End Sub

Public Sub TABLE_AfterPost()
  If InclusaoForaDoPadrao Then
    WriteAudit("|", HandleOfTable("SAM_PLANO_MODPRC_FX"), CurrentQuery.FieldByName("HANDLE").Value, LogInclusaoForaDoPadrao)
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim TipoCalculoPreco As String

  TipoCalculoPreco = VerificaPreco

  If TipoCalculoPreco <>"2" And _
      Not(CurrentQuery.FieldByName("PLANOTPDEP").IsNull)Then
    CanContinue = False
    bsShowMessage("A configuração do módulo NÃO permite que se informe o TipoDependente", "E")
  End If

  If TipoCalculoPreco = "3" And _
                        CurrentQuery.FieldByName("GRUPODEPENDENTE").IsNull Then
    CanContinue = False
    bsShowMessage("A configuração do módulo exige que se informe o grupo de dependente", "E")
  End If


  InclusaoForaDoPadrao = False

  If CurrentQuery.State = 3 Then
    Dim SQL As Object
    Dim vCompetencia As Date
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT DATAPRECO")
    SQL.Add("FROM SAM_PLANO_MODPRC")
    SQL.Add("WHERE HANDLE = :HPLANOMODPRC")
    SQL.ParamByName("HPLANOMODPRC").Value = RecordHandleOfTable("SAM_PLANO_MODPRC")
    SQL.Active = True

    vCompetencia = DateValue("1/" + _
                   Str(DatePart("m", SQL.FieldByName("DATAPRECO").AsDateTime)) + "/" + _
                   Str(DatePart("yyyy", SQL.FieldByName("DATAPRECO").AsDateTime)))

    SQL.Clear
    SQL.Add("SELECT C.VALORMINIMO, C.VALORCUSTO")
    SQL.Add("FROM SAM_PLANO_MOD A, SAM_MODULO_PRECO B, SAM_MODULO_PRECO_FX C")
    SQL.Add("WHERE A.HANDLE = :HPLANOMOD")
    SQL.Add("  AND B.MODULO = A.MODULO")
    SQL.Add("  AND B.COMPETENCIAINICIAL <= :COMPETENCIA")
    SQL.Add("  AND (B.COMPETENCIAFINAL IS NULL OR B.COMPETENCIAFINAL >= :COMPETENCIA)")
    SQL.Add("  AND C.MODULOPRECO = B.HANDLE")
    SQL.Add("  AND C.IDADEMAXIMA = (SELECT MIN(IDADEMAXIMA)")
    SQL.Add("                       FROM SAM_MODULO_PRECO_FX")
    SQL.Add("                       WHERE MODULOPRECO = B.HANDLE")
    SQL.Add("                         AND IDADEMAXIMA >= :IDADEMAXIMA)")
    SQL.ParamByName("HPLANOMOD").Value = RecordHandleOfTable("SAM_PLANO_MOD")
    SQL.ParamByName("IDADEMAXIMA").Value = CurrentQuery.FieldByName("IDADEMAXIMA").AsInteger
    SQL.ParamByName("COMPETENCIA").Value = vCompetencia
    SQL.Active = True

    If Not SQL.EOF Then
      If CurrentQuery.FieldByName("VALOR").AsFloat <SQL.FieldByName("VALORCUSTO").AsFloat Then
        If bsShowMessage("O valor informado está abaixo do valor de custo padrão do módulo. Deseja continuar?", "Q") = vbYes Then
          InclusaoForaDoPadrao = True
          LogInclusaoForaDoPadrao = "Inclusão de faixa abaixo do valor de custo padrão" + Chr(13) + "Custo: " + SQL.FieldByName("VALORCUSTO").AsString + "  Valor: " + CurrentQuery.FieldByName("VALOR").AsString
        Else
          CanContinue = False
          Set SQL = Nothing
          Exit Sub
        End If
      End If
      If CurrentQuery.FieldByName("VALOR").AsFloat <SQL.FieldByName("VALORMINIMO").AsFloat Then
      	If bsShowMessage("O valor está abaixo do valor mínimo padrão do módulo. Deseja continuar?", "Q") = vbYes Then
		  InclusaoForaDoPadrao = True
   	      LogInclusaoForaDoPadrao = "Inclusão de faixa abaixo do valor mínimo padrão" + Chr(13) + "Mínimo: " + SQL.FieldByName("VALORMINIMO").AsString + "  Valor: " + CurrentQuery.FieldByName("VALOR").AsString
    	Else
		  CanContinue = False
		  Set SQL = Nothing
      	  Exit Sub
        End If
      End If
    End If

    Set SQL = Nothing

  End If
End Sub

Public Function VerificaPreco As String
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT B.TIPOCALCULOPRECO FROM SAM_PLANO_MODPRC A, SAM_PLANO_MOD B")
  SQL.Add("WHERE A.HANDLE = :HANDLE")
  SQL.Add("  AND B.HANDLE = A.PLANOMODULO")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PLANOMODULOPRECO").Value
  SQL.Active = True
  VerificaPreco = SQL.FieldByName("TIPOCALCULOPRECO").AsString

  SQL.Active = False
  Set SQL = Nothing
End Function

