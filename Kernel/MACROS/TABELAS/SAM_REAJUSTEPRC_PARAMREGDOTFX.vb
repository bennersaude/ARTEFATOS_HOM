'HASH: 1900FD6B67CD581B4C644BBBAB63360F
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit
Dim v_EventoInicial As String
Dim v_EventoFinal As String


Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(False, EVENTOINICIAL.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
  '  End If
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  '  If Len(EVENTO.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(False, EVENTOFINAL.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
  '  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Q As Object
  Set Q = NewQuery

  Q.Active = False
  Q.Add("SELECT T1.ESTRUTURA EVENTOINICIAL,")
  Q.Add("       T2.ESTRUTURA EVENTOFINAL   ")
  Q.Add("  FROM SAM_TGE T1,                ")
  Q.Add("       SAM_TGE T2                 ")
  Q.Add(" WHERE T1.HANDLE = :EVENTOINICIAL ")
  Q.Add("   AND T2.HANDLE = :EVENTOFINAL   ")

  Q.ParamByName("EVENTOINICIAL").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
  Q.ParamByName("EVENTOFINAL").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
  Q.Active = True

  If Q.FieldByName("EVENTOINICIAL").AsString > Q.FieldByName("EVENTOFINAL").AsString Then
    bsShowMessage("Evento Inicial não pode ser maior que Evento Final !", "E")
    CanContinue = False
  End If

  Set Q = Nothing
End Sub

