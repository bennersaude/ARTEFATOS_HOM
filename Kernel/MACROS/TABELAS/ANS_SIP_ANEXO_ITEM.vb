'HASH: 507986506CE783446BB1AB7CA85DC3E1
'Macro: ANS_SIP_ANEXO_ITEM
'#Uses "*bsShowMessage"

Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object
  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If
  Set Obj = CreateBennerObject("BSANS001.uRotinas")
  Obj.ExcluirEventoClasse(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
End Sub

Public Sub BOTAOEXCLUIRGRAU_OnClick()
  Dim Obj As Object
  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If
  Set Obj = CreateBennerObject("BSANS001.uRotinas")
  Obj.ExcluirGrau(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
End Sub

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object
  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If
  Set Obj = CreateBennerObject("BSANS001.uRotinas")
  Obj.GerarEventoClasse(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
End Sub

Public Sub BOTAOGERARGRAU_OnClick()
  Dim Obj As Object
  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If
  Set Obj = CreateBennerObject("BSANS001.uRotinas")
  Obj.GerarGrau(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
End Sub

Public Sub BOTAOIMPORTARPARAMETROS_OnClick()
  'Dim Obj As Object
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  If CurrentQuery.State <> 1 Then
  	bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  Set vcContainer = NewContainer

  INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0076", _
					   "Importar Parâmetros",  _
					   CurrentQuery.FieldByName("HANDLE").AsInteger, _
					   200, _
					   300, _
					   False, _
					   vsMensagem, _
					   vcContainer)

  Set INTERFACE0002 = Nothing

  If Not WebMode Then
  	RefreshNodesWithTable("ANS_SIP_ANEXO_ITEM")
  End If
End Sub

Public Sub IDENTIFICARINTERNACAO_OnChange()
 If CurrentQuery.FieldByName("IDENTIFICARINTERNACAO").AsString = "S" Then
  INTERNACAO24HRS.Visible = True
 Else
  INTERNACAO24HRS.Visible = False
 End If
End Sub

Public Sub OBITO_OnChange()
  If CurrentQuery.FieldByName("OBITO").AsString = "S" Then
'    INTERNACAO.Visible = True
  Else
'    INTERNACAO.Visible = False
  End If
End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
Dim QTmp As Object
Set QTmp = NewQuery

QTmp.Active = False
QTmp.Clear
QTmp.Add("SELECT APROPRIACAO,ANEXO ")
QTmp.Add("  FROM ANS_SIP_ANEXO_ITEM ")
QTmp.Add(" WHERE HANDLE = :HND")
QTmp.ParamByName("HND").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
QTmp.Active = True
If QTmp.FieldByName("APROPRIACAO").AsString = "E" Then
  BOTAOGERAR.Caption = "Gerar Eventos"
  BOTAOEXCLUIR.Caption = "Excluir Eventos"
End If
If QTmp.FieldByName("APROPRIACAO").AsString = "C" Then
  BOTAOGERAR.Caption = "Gerar Classes"
  BOTAOEXCLUIR.Caption = "Excluir Classes"
End If

If CurrentQuery.FieldByName("OBITO").AsString = "S" Then
'  INTERNACAO.Visible = True
Else
'  INTERNACAO.Visible = False
End If

If CurrentQuery.FieldByName("ANEXO").AsString = "1" Then
	DESCRICAO.ReadOnly = True
Else
	DESCRICAO.ReadOnly = False
End If

If CurrentQuery.FieldByName("ANEXO").AsString <> "2" Then
  CONTAGEMQUANTIDADE.Visible = False
  CONTAGEMEXPOSTOS.Visible = True
Else
  CONTAGEMQUANTIDADE.Visible = False
  CONTAGEMEXPOSTOS.Visible = False
End If

If CurrentQuery.FieldByName("IDENTIFICARINTERNACAO").AsString = "S" Then
  INTERNACAO24HRS.Visible = True
Else
  INTERNACAO24HRS.Visible = False
End If

If CurrentQuery.FieldByName("ODONTOLOGICO").AsString = "S" Then
  BOTAOGERARGRAU.Visible   = True
  BOTAOEXCLUIRGRAU.Visible = True
Else
  BOTAOGERARGRAU.Visible   = False
  BOTAOEXCLUIRGRAU.Visible = False
End If

If CurrentQuery.FieldByName("ANEXO").AsString = "4" Then
  PRIMEIRAOCORRENCIAANO.Visible = True
  OBITO.Visible = False
  CONTAGEM.Visible = False
  CONTAGEMQUANTIDADE.Visible = False
  OBITOINTERNACAO.Visible = True
Else
  PRIMEIRAOCORRENCIAANO.Visible = False
  OBITO.Visible = False
  CONTAGEM.Visible = False
  CONTAGEMQUANTIDADE.Visible = False
  OBITOINTERNACAO.Visible = False
End If

ANEXO.Visible = False

End Sub

Public Sub TABLE_NewRecord()
Dim SQL As Object
   Set SQL = NewQuery
   SQL.Clear
   SQL.Add("SELECT ANEXO FROM ANS_SIP_ANEXO WHERE HANDLE = :HANDLE ")
   SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("ANS_SIP_ANEXO")
   SQL.Active = True
   CurrentQuery.FieldByName("ANEXO").AsInteger = SQL.FieldByName("ANEXO").AsInteger
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOIMPORTARPARAMETROS"
			BOTAOIMPORTARPARAMETROS_OnClick
	End Select
End Sub
