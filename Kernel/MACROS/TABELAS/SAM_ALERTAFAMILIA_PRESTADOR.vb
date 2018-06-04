'HASH: 2D57C8792AA6929A796DC2EE0F8BAE3F
'Macro: SAM_ALERTAFAMILIA_PRESTADOR
Option Explicit

'#Uses "*ProcuraPrestador"
'#Uses "*bsShowMessage"

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  '   If Len(PRESTADOR.Text) = 0 Then
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
  Q.Add("SELECT * FROM SAM_ALERTAFAMILIA WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  Q.Active = True

  CurrentQuery.RequestLive = Q.FieldByName("DATAFINAL").IsNull
  PRESTADOR .ReadOnly = Not Q.FieldByName("DATAFINAL").IsNull
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT USUARIO FROM SAM_ALERTAFAMILIA WHERE HANDLE=:HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  q1.Active = True
  If q1.FieldByName("USUARIO").AsInteger<>CurrentUser Then
    bsShowMessage("O usuário atual não tem permissão para alteração", "E")
    CanContinue = False
  End If
  q1.Active = False
  Set q1 = Nothing
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q1 As Object
  Dim Q2 As Object

  Set q1 = NewQuery
  q1.Add("SELECT USUARIO FROM SAM_ALERTAFAMILIA WHERE HANDLE=:HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  q1.Active = True
  If q1.FieldByName("USUARIO").AsInteger<>CurrentUser Then
    bsShowMessage("O usuário atual não tem permissão para alteração", "E")
    CanContinue = False
  End If
  q1.Active = False
  Set q1 = Nothing
  '**************************************************************************************************************
  '************ AlteraþÒo Para nÒo deixar cadastrar o mesmo prestador mais de uma vez no mesmo alerta  **********
  '**************************************************************************************************************
  Set Q2 = NewQuery
  Q2.Add("SELECT HANDLE                        ")
  Q2.Add("  FROM SAM_ALERTAFAMILIA_PRESTADOR   ")
  Q2.Add(" WHERE PRESTADOR = :PRESTADOR        ")
  Q2.Add("   AND FAMILIAALERTA = :FAMILIAALERTA")
  Q2.ParamByName("FAMILIAALERTA").AsInteger = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  Q2.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  Q2.Active = True
  If Not Q2.EOF Then
    bsShowMessage("Este prestador já está cadastrado neste alerta!", "E")
    CanContinue = False
  End If
  Q2.Active = False
  Set Q2 = Nothing
  '************************** Fim da ALteraþÒo ******************************************************************

  '******************************************************************************************************
  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_ALERTAFAMILIA WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("FAMILIAALERTA").AsInteger
  Q.Active = True

  If Not Q.FieldByName("DATAFINAL").IsNull Then
    bsShowMessage("Cadastro não permitido, pois a vigência do alerta está fechada", "E")
    CanContinue = False
    Exit Sub
  End If
  '******************************************************************************************************

End Sub
