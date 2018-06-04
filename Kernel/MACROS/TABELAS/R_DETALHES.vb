'HASH: 86357C691E8B3146C717D40B9A270541
'#Uses "*bsShowMessage"

Option Explicit
Dim bClearFields As Boolean

 
Public Sub CRIARCOMOGENERICO_OnClick() 
Dim HandleDetalhe As Integer, HandleCampo As Integer, ORDEM As Integer 
Dim ListaParam As String, TipoMsg As String 
Dim Q As Object, QAux As Object 
  Select Case CurrentQuery.FieldByName("TIPO").AsInteger 
    Case 2 
      TipoMsg = "cabeçalho" 
    Case 3 
      TipoMsg = "rodapé" 
    Case 4 
      TipoMsg = "sumário" 
  End Select 
 
  If (bsShowMessage("Será criado um " + TipoMsg + " genérico baseado no ítem selecionado." + Chr(10) + "Deseja continuar?", "Q") = vbYes) Then
 	HandleDetalhe = CopyRecord("R_DETALHES", CurrentQuery.FieldByName("HANDLE").AsInteger,"")
 
 	Set Q = NewQuery
  	Q.Add("UPDATE R_DETALHES SET RELATORIO = :RELATORIO WHERE HANDLE = :HANDLE")
  	Q.ParamByName("RELATORIO").DataType = ftInteger
  	Q.ParamByName("RELATORIO").Clear
  	Q.ParamByName("HANDLE").Value = HandleDetalhe
  	Q.ExecSQL
 
  	Set QAux = NewQuery
  	QAux.Add("SELECT MAX(ORDEM) ORDEM FROM R_DETALHECAMPOS WHERE DETALHE = :DETALHE")
  	QAux.ParamByName("DETALHE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  	QAux.Active = True
 
  	If Not QAux.FieldByName("ORDEM").IsNull Then
	    ORDEM = QAux.FieldByName("ORDEM").AsInteger
  	End If
 
  	QAux.Clear
  	QAux.Add("SELECT HANDLE, ORDEM FROM R_DETALHECAMPOS WHERE DETALHE = :DETALHE")
  	QAux.ParamByName("DETALHE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  	QAux.Active = True
 
  	While Not QAux.EOF
	    ORDEM = ORDEM + 10
	    ListaParam = "ORDEM=" + Format(ORDEM, "000")
	    HandleCampo = CopyRecord("R_DETALHECAMPOS", QAux.FieldByName("HANDLE").AsInteger, ListaParam)

	    Q.Add("UPDATE R_DETALHECAMPOS SET DETALHE = :DETALHE, ORDEM = :ORDEM WHERE HANDLE = :HANDLE")
	    Q.ParamByName("DETALHE").AsInteger = HandleDetalhe
	    Q.ParamByName("ORDEM").AsString = QAux.FieldByName("ORDEM").AsString
	    Q.ParamByName("HANDLE").AsInteger = HandleCampo
	    Q.ExecSQL
 
 	   QAux.Next
  	Wend
 
  	Set QAux = Nothing
  	Set Q = Nothing
  End If
End Sub 
 
Public Sub EXPORTAR_OnClick() 
  Dim obj As Object 
  Set obj = CreateBennerObject("CS.ReportFunctions") 
 
  obj.ExportDetail(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("NOME").AsString + ".DRP") 
 
  Set obj = Nothing 
End Sub 
 
Public Sub TABLE_AfterInsert()
	If VisibleMode Then
  		CurrentQuery.FieldByName("TIPO").Value = NodeInternalCode
	ElseIf WebMode Then
		Select Case WebVisionCode
			Case "V_R_DETALHES_153" 'Tipo: 1 - Detalhe
  				CurrentQuery.FieldByName("TIPO").Value = 1
  			Case "V_R_DETALHES_180" 'Tipo: 2 - Cabeçalho
				CurrentQuery.FieldByName("TIPO").Value = 2
			Case "V_R_DETALHES_181" 'Tipo: 3 - Rodapés
				CurrentQuery.FieldByName("TIPO").Value = 3
			Case "V_R_DETALHES_182" 'Tipo: 4 - Sumários
				CurrentQuery.FieldByName("TIPO").Value = 4
			Case "V_R_DETALHES_340" 'Tipo: 5 - Título
				CurrentQuery.FieldByName("TIPO").Value = 5
		End Select
	End If
End Sub 
Public Sub TABLE_AfterPost() 
Dim Q As Object 
If bClearFields Then 
  Set Q = NewQuery 
  Q.Add("UPDATE R_DETALHECAMPOS SET TABELA = :TABELA, CAMPO = NULL WHERE DETALHE = :DETALHE") 
  Q.ParamByName("DETALHE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
  Q.ParamByName("TABELA").AsInteger = CurrentQuery.FieldByName("TABELA").AsInteger 
  Q.ExecSQL 
  Set Q = Nothing 
End If 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  If CurrentQuery.FieldByName("RELATORIO").AsInteger = 0 Then 
    CRIARCOMOGENERICO.Visible = False 
  Else 
    CRIARCOMOGENERICO.Visible = True 
  End If 
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Q As Object


  CurrentQuery.UpdateRecord

  If  (VisibleMode And (CurrentQuery.FieldByName("TIPO").AsInteger <> NodeInternalCode)) _
     And (CurrentQuery.FieldByName("HANDLE").AsInteger > 0)  Then

    Set Q = NewQuery
    Q.Add("SELECT COUNT(HANDLE) NRECS FROM R_DETALHEDETALHES WHERE DETALHE = " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger ))
    Q.Active = True
    If Q.FieldByName("NRECS").AsInteger > 0 Then
      bsShowMessage("Operação não permitida para detalhes com sub-detalhes!", "I")
      TABLE_AfterInsert
    Else
      Q.Active = False
      Q.Clear
      Q.Add("SELECT COUNT(HANDLE) NRECS FROM R_QUEBRASDETALHE WHERE DETALHE = " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger))
      Q.Active = True
      If Q.FieldByName("NRECS").AsInteger > 0 Then
        bsShowMessage("Operação não permitida para detalhes com quebras!", "I")
        TABLE_AfterInsert
      End If
    End If
    Q.Active = False
    Set Q = Nothing
  End If

  If WebMode Then
  	Select Case WebVisionCode
  		Case "V_R_DETALHES_153" 'Tipo: 1 - Detalhe
  			If (CurrentQuery.FieldByName("TIPO").AsInteger <> 1 And CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
  				VerificaRelatorio
  			End If
  		Case "V_R_DETALHES_180" 'Tipo: 2 - Cabeçalho
			If (CurrentQuery.FieldByName("TIPO").AsInteger <> 2 And CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
  				VerificaRelatorio
  			End If
		Case "V_R_DETALHES_181" 'Tipo: 3 - Rodapés
			If (CurrentQuery.FieldByName("TIPO").AsInteger <> 3 And CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
  				VerificaRelatorio
  			End If
		Case "V_R_DETALHES_182" 'Tipo: 4 - Sumários
			If (CurrentQuery.FieldByName("TIPO").AsInteger <> 4 And CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
  				VerificaRelatorio
  			End If
		Case "V_R_DETALHES_340" 'Tipo: 5 - Título
			If (CurrentQuery.FieldByName("TIPO").AsInteger <> 5 And CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
  				VerificaRelatorio
  			End If
  	End Select
  End If

If Not CurrentQuery.FieldByName("RELATORIO").IsNull Then 
Set Q = NewQuery 
Q.Add("SELECT COUNT(HANDLE) NRECS FROM R_DETALHES WHERE RELATORIO = :RELATORIO AND NOME = :NOME AND HANDLE <> :HANDLE") 
Q.ParamByName("RELATORIO").AsInteger = CurrentQuery.FieldByName("RELATORIO").AsInteger 
Q.ParamByName("NOME").AsString = CurrentQuery.FieldByName("NOME").AsString 
  Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
  Q.Active = True 
  If Q.FieldByName("NRECS").AsInteger > 0 Then 
    bsShowMessage("Já existe um detalhe com este nome.","E")
    CanContinue = False 
  End If 
  Q.Active = False 
  Set Q = Nothing 
End If 
 
bClearFields = False 
If Not CurrentQuery.FieldByName("TABELA").IsNull Then 
  Set Q = NewQuery 
  Q.Add("SELECT COUNT(HANDLE) NRECS FROM R_DETALHECAMPOS WHERE DETALHE = :DETALHE AND (TABELA <> :TABELA OR TABELA IS NULL)") 
  Q.ParamByName("DETALHE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger 
  Q.ParamByName("TABELA").AsInteger = CurrentQuery.FieldByName("TABELA").AsInteger 
  Q.Active = True 
  If Q.FieldByName("NRECS").AsInteger > 0 Then 
    If bsShowMessage("Campos deste detalhe, que pertençam a outra tabela, deverão ser informados novamente." + Chr(13)+Chr(10) +"Confirma a operação?", "Q" ) = vbYes Then
      bClearFields = True 
    Else 
      CanContinue = False 
    End If 
  End If 
  Q.Active = False 
  Set Q = Nothing 
End If 
 
End Sub 
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case CRIARCOMOGENERICO
			CRIARCOMOGENERICO_OnClick
		Case EXPORTAR
			EXPORTAR_OnClick
		Case TIRARFONTES
			TIRARFONTES_OnClick
	End Select
End Sub

 
Public Sub TIRARFONTES_OnClick() 
Dim Sql, Sql2 
  If bsShowMessage("Deseja limpar todas a fontes desse detalhe e seus filhos? ", "Q") = vbYes Then
	  If InTransaction Then
	    Rollback
	  End If
  	StartTransaction
  	On Error GoTo ProcessaErro
  		Sql.Active = False
	    Sql.Clear
	    Sql.Add "UPDATE R_DETALHES SET FONTE = NULL WHERE HANDLE = " + CurrentQuery.FieldByName("HANDLE").AsString
  		Sql.ExecSQL
  		Sql.Active = False
	    Sql.Clear
	    Sql.Add "SELECT HANDLE FROM R_DETALHES WHERE HANDLE = " + CurrentQuery.FieldByName("HANDLE").AsString
  		Sql.Active = True
  	While Not Sql.EOF
      	Sql2.Active = False
      	Sql2.Clear
      	Sql2.Add "UPDATE R_DETALHEDETALHES SET FONTELEGENDAS = NULL WHERE DETALHE = " + Sql.FieldByName("HANDLE").AsString
    	Sql2.ExecSQL
      	Sql2.Active = False
      	Sql2.Clear
      	Sql2.Add "UPDATE R_DETALHECAMPOS SET FONTE = NULL WHERE DETALHE = " + Sql.FieldByName("HANDLE").AsString
    	Sql2.ExecSQL
    	Sql.Next
    Wend
  	Set Sql = Nothing
  	Set Sql2 = Nothing
  	Commit
  End If
  Exit Sub 
ProcessaErro: 
  Rollback 
End Sub 

Public Sub VerificaRelatorio()
Dim Q As Object

	Set Q = NewQuery
    Q.Add("SELECT COUNT(HANDLE) NRECS FROM R_DETALHEDETALHES WHERE DETALHE = " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger ))
    Q.Active = True
    If Q.FieldByName("NRECS").AsInteger > 0 Then
      bsShowMessage("Operação não permitida para detalhes com sub-detalhes!", "I")
      TABLE_AfterInsert
    Else
      Q.Active = False
      Q.Clear
      Q.Add("SELECT COUNT(HANDLE) NRECS FROM R_QUEBRASDETALHE WHERE DETALHE = " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger))
      Q.Active = True
      If Q.FieldByName("NRECS").AsInteger > 0 Then
        bsShowMessage("Operação não permitida para detalhes com quebras!", "I")
        TABLE_AfterInsert
      End If
    End If
    Q.Active = False
    Set Q = Nothing

End Sub
