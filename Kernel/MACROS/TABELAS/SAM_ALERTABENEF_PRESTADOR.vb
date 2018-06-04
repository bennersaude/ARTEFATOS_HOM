'HASH: F6E2C3F696CD49C83944F292E3BE40A7
'Macro: SAM_ALERTABENEF_PRESTADOR
'#uses "*bsShowMessage"
Option Explicit

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  '#Uses "*ProcuraPrestador"
  '  If Len(PRESTADOR.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("N", "T", PRESTADOR.Text) ' pelo Nome e Todos
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  ' End If
End Sub


Public Sub TABLE_AfterScroll()
  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_ALERTABENEF WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("BENEFICIARIOALERTA").AsInteger
  Q.Active = True

  CurrentQuery.RequestLive = Q.FieldByName("DATAFINAL").IsNull

  	If (Q.FieldByName("DATAFINAL").IsNull Or ServerDate < Q.FieldByName("DATAFINAL").AsDateTime) Then
  		CurrentQuery.RequestLive = True
		PRESTADOR.ReadOnly = False
  	Else
		PRESTADOR.ReadOnly = True
  		bsShowMessage("Data final de vigência já encerrada!", "I")
  	End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT USUARIO FROM SAM_ALERTABENEF WHERE HANDLE=:HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("BENEFICIARIOALERTA").AsInteger
  q1.Active = True
  If q1.FieldByName("USUARIO").AsInteger<>CurrentUser Then
    bsShowMessage("O usuário atual não tem permissão para alteração","E")
    CanContinue = False
  End If
  q1.Active = False
  Set q1 = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q1 As Object
  Dim Q2 As Object
  Set q1 = NewQuery

  q1.Add("SELECT USUARIO FROM SAM_ALERTABENEF WHERE HANDLE=:HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("BENEFICIARIOALERTA").AsInteger
  q1.Active = True
  If q1.FieldByName("USUARIO").AsInteger<>CurrentUser Then
    bsShowMessage("O usuário atual não tem permissão para alteração","E")
    CanContinue = False
  End If
  q1.Active = False
  Set q1 = Nothing
  '**************************************************************************************************************
  '************ Alteração Para não deixar cadastrar o mesmo prestador mais de uma vez no mesmo alerta  **********
  '**************************************************************************************************************
  Set Q2 = NewQuery
  Q2.Add("SELECT HANDLE                         ")
  Q2.Add("  FROM SAM_ALERTABENEF_PRESTADOR      ")
  Q2.Add(" WHERE PRESTADOR = :PRESTADOR         ")
  Q2.Add("   AND BENEFICIARIOALERTA = :BENEFICIARIOALERTA ")
  Q2.ParamByName("BENEFICIARIOALERTA").AsInteger = CurrentQuery.FieldByName("BENEFICIARIOALERTA").AsInteger
  Q2.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  Q2.Active = True
  If Not Q2.EOF Then
    bsShowMessage("Este prestador já está cadastrado neste alerta!", "E")
    CanContinue = False
  End If
  Q2.Active = False
  Set Q2 = Nothing
  '************************** Fim da ALteração ******************************************************************



  '******************************************************************************************************
  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_ALERTABENEF WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("BENEFICIARIOALERTA").AsInteger
  Q.Active = True

  If (Not Q.FieldByName("DATAFINAL").IsNull) And (ServerDate > Q.FieldByName("DATAFINAL").AsDateTime) Then
    bsShowMessage("Cadastro não permitido, pois a vigência do alerta está fechada", "E")
    CanContinue = False
    Exit Sub
  End If
  '******************************************************************************************************


End Sub

