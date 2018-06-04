'HASH: 51A5A4E719D025B26F3A127FF3145D0F
'#Uses "*bsShowMessage"
Option Explicit
Dim ChangingLer As Boolean 
 
Sub VerificaLer 
  CurrentQuery.UpdateRecord 
  If (Not ChangingLer) And (CurrentQuery.FieldByName("LER").AsString = "N") And ((CurrentQuery.FieldByName("ALTERAR").AsString = "S") Or (CurrentQuery.FieldByName("EXCLUIR").AsString = "S") Or (CurrentQuery.FieldByName("INCLUIR").AsString = "S")) Then
    CurrentQuery.FieldByName("LER").AsString = "S" 
  End If 
End Sub 
 
Public Sub BOTAOLIBERARGLOSA_OnClick()
  Dim SQL1 As Object
  Dim SQL2 As Object
  Dim SQLUSUARIO As Object

  If bsShowMessage("Confirma a liberação de todas as glosas para o usuário ?", "Q") = vbYes Then
	Set SQL1 = NewQuery
	Set SQL2 = NewQuery
	Set SQLUSUARIO = NewQuery

    SQL1.Clear
	SQL1.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA")
	SQL1.Add("WHERE HANDLE NOT IN")
	SQL1.Add("(SELECT MOTIVOGLOSA FROM SAM_GRUPO_MOTIVOGLOSA WHERE GRUPO = :GRUPO)")
	SQL1.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger ' RecordHandleOfTable("Z_GRUPOS")
	SQL1.Active=True

	While Not SQL1.EOF
	  SQL2.Clear
	  SQL2.Add("INSERT INTO SAM_GRUPO_MOTIVOGLOSA (HANDLE, GRUPO, MOTIVOGLOSA) VALUES (:HANDLE,:GRUPO,:MOTIVOGLOSA)")
	  SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_GRUPO_MOTIVOGLOSA")
	  SQL2.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger ' RecordHandleOfTable("Z_GRUPOS")
	  SQL2.ParamByName("MOTIVOGLOSA").Value = SQL1.FieldByName("HANDLE").AsInteger
	  SQL2.ExecSQL
	  SQL1.Next
	Wend
	RefreshNodesWithTable"SAM_GRUPO_MOTIVOGLOSA"


    SQLUSUARIO.Clear
    SQLUSUARIO.Add("SELECT HANDLE FROM Z_GRUPOUSUARIOS WHERE GRUPO = :GRUPO")
    SQLUSUARIO.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger ' RecordHandleOfTable("Z_GRUPOS")
    SQLUSUARIO.Active = True

    While Not SQLUSUARIO.EOF
      Set SQL1 = NewQuery
	  Set SQL2 = NewQuery


      SQL1.Clear
      SQL1.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA")
      SQL1.Add("WHERE HANDLE NOT IN")
      SQL1.Add("(SELECT MOTIVOGLOSA FROM SAM_USUARIO_MOTIVOGLOSA WHERE USUARIO = :USUARIO)")
      SQL1.Add("AND HANDLE IN (SELECT MOTIVOGLOSA FROM SAM_GRUPO_MOTIVOGLOSA WHERE GRUPO = :GRUPO)")
      SQL1.ParamByName("USUARIO").Value = SQLUSUARIO.FieldByName("HANDLE").AsInteger
      SQL1.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger ' RecordHandleOfTable("Z_GRUPOS")
      SQL1.Active=True

      While Not SQL1.EOF
        SQL2.Clear
        SQL2.Add("INSERT INTO SAM_USUARIO_MOTIVOGLOSA (HANDLE, USUARIO, MOTIVOGLOSA) VALUES (:HANDLE,:USUARIO,:MOTIVOGLOSA)")
        SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVOGLOSA")
        SQL2.ParamByName("USUARIO").Value = SQLUSUARIO.FieldByName("HANDLE").AsInteger
        SQL2.ParamByName("MOTIVOGLOSA").Value = SQL1.FieldByName("HANDLE").AsInteger
        SQL2.ExecSQL
        SQL1.Next
      Wend

     SQLUSUARIO.Next

    Wend
  End If
End Sub

Public Sub BOTAOLIBERARNEGACAO_OnClick()
  Dim SQL1 As Object
  Dim SQL2 As Object
  Dim SQLUSUARIO As Object

  If bsShowMessage("Confirma a liberação de todas as negações para o usuário ?" , "Q") = vbYes Then
	Set SQL1 = NewQuery
	Set SQL2 = NewQuery
	Set SQLUSUARIO = NewQuery

    SQL1.Clear
	SQL1.Add("SELECT HANDLE FROM SAM_MOTIVONEGACAO")
	SQL1.Add("WHERE HANDLE NOT IN")
	SQL1.Add("(SELECT MOTIVONEGACAO FROM SAM_GRUPO_MOTIVONEGACAO WHERE GRUPO = :GRUPO)")
	SQL1.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger ' RecordHandleOfTable("Z_GRUPOS")
	SQL1.Active=True

	While Not SQL1.EOF
	  SQL2.Clear
	  SQL2.Add("INSERT INTO SAM_GRUPO_MOTIVONEGACAO (HANDLE, GRUPO, MOTIVONEGACAO) VALUES (:HANDLE,:GRUPO,:MOTIVONEGACAO)")
	  SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_GRUPO_MOTIVONEGACAO")
	  SQL2.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger ' RecordHandleOfTable("Z_GRUPOS")
	  SQL2.ParamByName("MOTIVONEGACAO").Value = SQL1.FieldByName("HANDLE").AsInteger
	  SQL2.ExecSQL
	  SQL1.Next
	Wend
	RefreshNodesWithTable"SAM_GRUPO_MOTIVONEGACAO"

    SQLUSUARIO.Clear
	SQLUSUARIO.Add("SELECT HANDLE FROM Z_GRUPOUSUARIOS WHERE GRUPO = :GRUPO")
	SQLUSUARIO.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger ' RecordHandleOfTable("Z_GRUPOS")
	SQLUSUARIO.Active = True

    While Not SQLUSUARIO.EOF
	  Set SQL1 = NewQuery
	  Set SQL2 = NewQuery


      SQL1.Clear
	  SQL1.Add("SELECT HANDLE FROM SAM_MOTIVONEGACAO")
	  SQL1.Add("WHERE HANDLE NOT IN")
	  SQL1.Add("(SELECT MOTIVONEGACAO FROM SAM_USUARIO_MOTIVONEGACAO WHERE USUARIO = :USUARIO)")
	  SQL1.Add("AND HANDLE IN (SELECT MOTIVONEGACAO FROM SAM_GRUPO_MOTIVONEGACAO WHERE GRUPO = :GRUPO)")
	  SQL1.ParamByName("USUARIO").Value = SQLUSUARIO.FieldByName("HANDLE").AsInteger
	  SQL1.ParamByName("GRUPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger ' RecordHandleOfTable("Z_GRUPOS")
	  SQL1.Active=True

	  While Not SQL1.EOF
	    SQL2.Clear
	    SQL2.Add("INSERT INTO SAM_USUARIO_MOTIVONEGACAO (HANDLE, USUARIO, MOTIVONEGACAO) VALUES (:HANDLE,:USUARIO,:MOTIVONEGACAO)")
	    SQL2.ParamByName("HANDLE").Value = NewHandle("SAM_USUARIO_MOTIVONEGACAO")
	    SQL2.ParamByName("USUARIO").Value = SQLUSUARIO.FieldByName("HANDLE").AsInteger
	    SQL2.ParamByName("MOTIVONEGACAO").Value = SQL1.FieldByName("HANDLE").AsInteger
	    SQL2.ExecSQL
	    SQL1.Next
	  Wend

      SQLUSUARIO.Next
    Wend
  End If
End Sub

Public Sub COPIARGRUPO_OnClick()
  Dim ObjCopy As Object 
    Set ObjCopy = CreateBennerObject("CS.Security") 
    ObjCopy.CopySecurityGroup(CurrentQuery.FieldByName("HANDLE").AsInteger, "NovoGrupo", CurrentSystem) 
    Set ObjCopy = Nothing 
End Sub 
 
Public Sub LER_OnChange() 
  ChangingLer = True 
  CurrentQuery.UpdateRecord 
 
  If (CurrentQuery.FieldByName("LER").AsString = "N") Then 
    CurrentQuery.FieldByName("ALTERAR").AsString = "N" 
    CurrentQuery.FieldByName("EXCLUIR").AsString = "N" 
    CurrentQuery.FieldByName("INCLUIR").AsString = "N" 
  End If 
  ChangingLer = False 
End Sub 
 
Public Sub INCLUIR_OnChange() 
  VerificaLer 
End Sub 
 
Public Sub ALTERAR_OnChange() 
  VerificaLer 
End Sub 
 
Public Sub EXCLUIR_OnChange() 
  VerificaLer 
End Sub 
 
Public Sub TABLE_AfterScroll() 
  ChangingLer = False
  TIPOAUTORIZACAOPADRAO.LocalWhere = "(DATAINICIAL <= " + CurrentSystem.SQLDate(CurrentSystem.ServerDate) + ") AND (DATAFINAL Is Null)" 'SMS 81867 - Débora Rebello - 21/05/2007
  If VisibleMode Then 'SMS 98111 - Barbosa - 19/06/2008
    BOTAOLIBERARGLOSA.Visible = False
    BOTAOLIBERARNEGACAO.Visible = False
  End If
End Sub 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("CRITICAINCPRESTADOR").AsString = "S" Then
		Dim qPrestador As BPesquisa
		Set qPrestador = NewQuery
		qPrestador.Add("SELECT COUNT(1) QTD                                            ")
		qPrestador.Add("  FROM SAM_PRESTADOR P                                         ")
		qPrestador.Add("  JOIN SAM_CATEGORIA_PRESTADOR CP ON (P.CATEGORIA = CP.HANDLE) ")
		qPrestador.Add("  JOIN Z_GRUPOUSUARIOS U ON (U.CPF = P.CPFCNPJ)                ")
		qPrestador.Add(" WHERE CP.BLOQUEARINCLUSAOBENEF = 'S'                          ")
		qPrestador.Add("   AND U.INATIVO = 'N'                                         ")
		qPrestador.Add("   AND U.GRUPO = :GRUPO                                        ")

		qPrestador.ParamByName("GRUPO").AsString = CurrentQuery.FieldByName("HANDLE").AsString
		qPrestador.Active = True
		qPrestador.First

	    If qPrestador.FieldByName("QTD").AsInteger > 0 Then
	        If (bsShowMessage("Existe usuário do grupo de segurança com o mesmo CPF de prestador. Deseja continuar?", "Q") = vbNo) Then
	            If VisibleMode Then
	                CanContinue = False
	            End If
	        End If
	    End If
    End If
End Sub

Public Sub TABLE_BeforeScroll()
	BOTAOEXCLUIR.Visible = False
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
    Case "BOTAOLIBERARGLOSA"
      BOTAOLIBERARGLOSA_OnClick
    Case "BOTAOLIBERARNEGACAO"
      BOTAOLIBERARNEGACAO_OnClick
  End Select
End Sub
