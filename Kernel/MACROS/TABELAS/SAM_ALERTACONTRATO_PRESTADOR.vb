'HASH: AC32E89E1D91A53616865C9235F02C52
'Macro: SAM_ALERTACONTRATO_PRESTADOR
Option Explicit

'#Uses "*ProcuraPrestador"
'#Uses "*bsShowMessage"

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  '  If Len(PRESTADOR.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("N", "T", PRESTADOR.Text) ' pelo Nome e Todos
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  '  End If

End Sub

Public Sub TABLE_AfterScroll()

  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_ALERTACONTRATO WHERE HANDLE = :CONTRATOALERTA")
  Q.ParamByName("CONTRATOALERTA").Value = CurrentQuery.FieldByName("CONTRATOALERTA").AsInteger
  Q.Active = True

  CurrentQuery.RequestLive = Q.FieldByName("DATAFINAL").IsNull
  PRESTADOR .ReadOnly = Not Q.FieldByName("DATAFINAL").IsNull

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Q1 As Object
  Set Q1 = NewQuery
  Q1.Add("SELECT USUARIO FROM SAM_ALERTACONTRATO WHERE HANDLE=:HANDLE")
  Q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CONTRATOALERTA").AsInteger
  Q1.Active = True
  If Q1.FieldByName("USUARIO").AsInteger<>CurrentUser Then
  	bsShowMessage("O usuário atual não tem permissão para alteração", "E")
  	CanContinue = False
  End If
  Q1.Active = False
  Set Q1 = Nothing
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Q1 As Object

  Set Q1 = NewQuery
  Q1.Add("SELECT USUARIO FROM SAM_ALERTACONTRATO WHERE HANDLE=:HANDLE")
  Q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CONTRATOALERTA").AsInteger
  Q1.Active = True
  If Q1.FieldByName("USUARIO").AsInteger<>CurrentUser Then
    bsShowMessage("O usuário atual não tem permissão para alteração", "E")
    CanContinue = False
  End If
  Q1.Active = False
  Set Q1 = Nothing
  '**************************************************************************************************************
  '************ Alteração Para não deixar cadastrar o mesmo prestador mais de uma vez no mesmo alerta  **********
  '**************************************************************************************************************
  Dim Q2 As Object
  Set Q2 = NewQuery
  Q2.Add("SELECT HANDLE                           ")
  Q2.Add("  FROM SAM_ALERTACONTRATO_PRESTADOR     ")
  Q2.Add(" WHERE PRESTADOR = :PRESTADOR           ")
  Q2.Add("   AND CONTRATOALERTA = :CONTRATOALERTA ")
  Q2.ParamByName("CONTRATOALERTA").AsInteger = CurrentQuery.FieldByName("CONTRATOALERTA").AsInteger
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
  Q.Add("SELECT * FROM SAM_ALERTACONTRATO WHERE HANDLE = :CONTRATOALERTA")
  Q.ParamByName("CONTRATOALERTA").Value = CurrentQuery.FieldByName("CONTRATOALERTA").AsInteger
  Q.Active = True

  If Not Q.FieldByName("DATAFINAL").IsNull Then
    bsShowMessage("Cadastro não permitido, pois a vigência do alerta está fechada", "E")
    CanContinue = False
    Exit Sub
  End If
  '******************************************************************************************************



End Sub
