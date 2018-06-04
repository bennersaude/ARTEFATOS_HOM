'HASH: FC4AFE9889AACAE32731DBEDB6DF31CE

'Macro: SAM_TGESUBSTITUTO

Option Explicit
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

End Sub

Public Sub TABLE_AfterScroll()
  If (WebMode) Then
    EVENTO.WebLocalWhere = " A.ULTIMONIVEL = 'S' "
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim q1 As Object
  Set q1 = NewQuery

  'Verifica se existe pelo menos o Grau Principal cadastrado no evento
  q1.Active = False
  q1.Clear
  q1.Add("SELECT STG.GRAU     GRAUPRINCIPAL")
  q1.Add("  FROM SAM_TGE_GRAU STG")
  q1.Add("  JOIN SAM_TGE      ST  ON ST.HANDLE = STG.EVENTO")
  q1.Add(" WHERE STG.GRAUPRINCIPAL = 'S'")
  q1.Add("   AND ST.HANDLE = :EVENTO")
  
  q1.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  q1.Active = True

  If q1.EOF Then
    bsShowMessage("O Evento Substituto precisa ter ao menos o Grau Principal cadastrado", "E")
    q1.Active = False
    Set q1 = Nothing
    CanContinue = False
    Exit Sub
  End If

  ' SMS 25690 - Douglas
  ' Não permite incluir um evento Inativo como Substituto
  q1.Active = False
  q1.Clear
  q1.Add("SELECT INATIVO FROM SAM_TGE WHERE HANDLE = :EVENTO")
  q1.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  q1.Active = True

  If q1.FieldByName("INATIVO").AsString = "S" Then
    bsShowMessage("Não é possível definir um evento Inativo como substituto.", "E")
    q1.Active = False
    Set q1 = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set q1 = Nothing

End Sub

