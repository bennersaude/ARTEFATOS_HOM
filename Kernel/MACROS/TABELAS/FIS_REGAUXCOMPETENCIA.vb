'HASH: 1B891CA9287E6BFB56E801E952D73434
'Macro: FIS_REGAUXCOMPETENCIA
Option Explicit
'#Uses "*PrimeiroDiaCompetencia"
'#Uses "*bsShowMessage"

Public Sub BOTAOAJUSTA_OnClick()
	Dim spProc As BStoredProc
    Set spProc = NewStoredProc

    spProc.Name = "Ajusta_UF_RA_03"
    spProc.AddParam("p_handle", ptInput, ftInteger)
    spProc.ParamByName("p_handle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    spProc.ExecProc

    Set spProc = Nothing
    Set spProc = NewStoredProc

    spProc.Name = "Ajusta_UF_RA_06"
    spProc.AddParam("p_handle", ptInput, ftInteger)
    spProc.ParamByName("p_handle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    spProc.ExecProc

    Set spProc = Nothing

    bsShowMessage("Processo concluído com sucesso!", "I")
End Sub

Public Sub BOTAOCANCELAR_OnClick()

  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  Set vcContainer = NewContainer

  INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0074", _
					   "Cancelar Registros Auxiliares",  _
					   CurrentQuery.FieldByName("HANDLE").AsInteger, _
					   225, _
					   500, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  Set INTERFACE0002 = Nothing

End Sub

Public Sub BOTAOPROCESSAR_OnClick()

  If CurrentQuery.State = 2 Then
    bsShowMessage("O registro não pode estar e edição.", "E")
    Exit Sub
  End If

  Dim Query As Object

  If CurrentQuery.FieldByName("PROVISORIO").AsString = "N" Then
    Set Query = NewQuery

    Query.Clear
    Query.Add("SELECT PERIODOFATCONINICIAL, CONTABILIZA FROM SFN_PARAMETROSFIN")
    Query.Active = True

    If Query.FieldByName("CONTABILIZA").AsString = "S" Then

      If (CurrentQuery.FieldByName("COMPETENCIA").AsDateTime >= PrimeiroDiaCompetencia(Query.FieldByName("PERIODOFATCONINICIAL").AsDateTime)) Then
        bsShowMessage("Processar uma rotina de Registro Auxiliar com competência dentro do período contábil só é permitido quando o flag provisório estiver marcado.", "E")
        Set Query = Nothing
        Exit Sub
      End If
    End If

    Set Query = Nothing

  End If

  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  Set vcContainer = NewContainer

  INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0073", _
					   "Processar Registros Auxiliares",  _
					   CurrentQuery.FieldByName("HANDLE").AsInteger, _
					   225, _
					   500, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  Set INTERFACE0002 = Nothing


End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.State = 3 Then
    BOTAOPROCESSAR.Enabled = False
    BOTAOCANCELAR.Enabled = False
    BOTAOAJUSTA.Enabled = False
  Else
  	If CurrentQuery.FieldByName("SITUACAOPROCESSO").AsString = "5" Then
	  BOTAOPROCESSAR.Enabled = False
      BOTAOAJUSTA.Enabled = True
      PROVISORIO.ReadOnly = True
      BOTAOCANCELAR.Enabled = True
    Else
      BOTAOPROCESSAR.Enabled = True
      BOTAOAJUSTA.Enabled = True
      If CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
        PROVISORIO.ReadOnly = False
        BOTAOCANCELAR.Enabled = False
      Else
        PROVISORIO.ReadOnly = True
        BOTAOCANCELAR.Enabled = True
	  End If
    End If
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

 '------------------- SMS 90104 - Paulo Melo - 17/12/2007 - INICIO  -- Fazer o tratamento para nao dar erro de unique do banco, mas sim mostra uma mensagem.
 Dim qCOMPETENCIA As Object
 Set qCOMPETENCIA = NewQuery

 qCOMPETENCIA.Add("SELECT HANDLE")
 qCOMPETENCIA.Add("FROM FIS_REGAUXCOMPETENCIA")
 qCOMPETENCIA.Add("WHERE COMPETENCIA = :COMPETENCIA")
 qCOMPETENCIA.Add("AND HANDLE <> :HANDLE")
 qCOMPETENCIA.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
 qCOMPETENCIA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 qCOMPETENCIA.Active = True

 If Not qCOMPETENCIA.EOF Then
 	bsShowMessage("Não é possível gravar duas competências iguais", "E")
 	CanContinue = False
 End If

 Set qCOMPETENCIA = Nothing
'------------------- SMS 90104 - Paulo Melo - 17/12/2007 - INICIO

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
 		Case "BOTAOCANCELAR"
 			BOTAOCANCELAR_OnClick
 		Case "BOTAOPROCESSAR"
 			BOTAOPROCESSAR_OnClick
 		Case "BOTAOAJUSTA"
 			BOTAOAJUSTA_OnClick
	End Select
End Sub
