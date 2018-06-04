'HASH: C008234B9AE5E3B98B80E02214F8A803

'Macro: SAM_TGESUBSTITUTO_ITEM

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


Public Sub TABLE_AfterPost()

  If (WebMode) Then
    EVENTO.WebLocalWhere = " A.ULTIMONIVEL = 'S' "
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  ' SMS 25690 - Douglas

  Dim q1 As Object
  Dim q2 As Object
  Set q1 = NewQuery
  Set q2 = NewQuery


  'Verifica se esse evento já está na lista, se estiver ele aborta a inclusão
  q1.Active = False
  q1.Clear
  q1.Add("SELECT COUNT(HANDLE) REPETICAO")
  q1.Add("  FROM SAM_TGESUBSTITUTO_ITEM ")
  q1.Add(" WHERE EVENTO = :EVENTO ")
  q1.Add("   AND TGESUBSTITUTO = :SUBSTITUTO ")


  q1.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  q1.ParamByName("SUBSTITUTO").AsInteger = CurrentQuery.FieldByName("TGESUBSTITUTO").AsInteger
  q1.Active = True

  If q1.FieldByName("REPETICAO").AsInteger > 0 Then
    bsShowMessage("Este evento já esta definido nesta lista de substituição", "E")
    Set q1 = Nothing
    CanContinue = False
    Exit Sub
  End If


  'Verifica se o evento sendo incluído não é o próprio substituto
  'Seleciona o handle na tge do evento substituto
  q1.Active = False
  q1.Clear
  q1.Add("SELECT EVENTO ")
  q1.Add("  FROM SAM_TGESUBSTITUTO ")
  q1.Add(" WHERE HANDLE = :SUBSTITUTO ")

  q1.ParamByName("SUBSTITUTO").AsInteger = CurrentQuery.FieldByName("TGESUBSTITUTO").AsInteger
  q1.Active = True
  'Verifica se o evento sendo incluído não é o mesmo que o substituto
  If q1.FieldByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger Then
    bsShowMessage("O evento não pode ser incluído em seu próprio conjunto de substituição.", "E")
    Set q1 = Nothing
    CanContinue = False
    Exit Sub
  End If


  'Verfica se o grau principal dos que estão sendo substituídos é o mesmo do substituto
  q1.Active = False
  q1.Clear
  q1.Add("SELECT STG.GRAU     GRAUPRINCIPAL")
  q1.Add("  FROM SAM_TGE_GRAU STG")
  q1.Add("  JOIN SAM_TGE      ST  ON ST.HANDLE = STG.EVENTO")
  q1.Add(" WHERE STG.GRAUPRINCIPAL = 'S'")
  q1.Add("   AND ST.HANDLE = (SELECT EVENTO ")
  q1.Add("                      FROM SAM_TGESUBSTITUTO ")
  q1.Add("                     WHERE HANDLE=:TGESUBSTITUTO)")

  q1.ParamByName("TGESUBSTITUTO").AsInteger = CurrentQuery.FieldByName("TGESUBSTITUTO").AsInteger
  q1.Active = True


  q2.Active = False
  q2.Clear
  q2.Add("SELECT STG.GRAU     GRAUPRINCIPAL")
  q2.Add("  FROM SAM_TGE_GRAU STG")
  q2.Add("  JOIN SAM_TGE      ST  ON ST.HANDLE = STG.EVENTO")
  q2.Add(" WHERE STG.GRAUPRINCIPAL = 'S'")
  q2.Add("   AND ST.HANDLE = :SUBSTITUIDO")

  q2.ParamByName("SUBSTITUIDO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  q2.Active = True

  If (Not q1.EOF) And (Not q2.EOF) Then
    If q1.FieldByName("GRAUPRINCIPAL").AsInteger <> q2.FieldByName("GRAUPRINCIPAL").AsInteger Then
      bsShowMessage("Este evento não pode ser substuído nesta lista. O Grau Principal do evento substituto e dos eventos a serem substituídos devem ser idênticos.", "E")
      q1.Active = False
      q2.Active = False
      Set q1 = Nothing
      Set q2 = Nothing
      CanContinue = False
      Exit Sub
    End If
  Else
    bsShowMessage("Este evento não pode ser substuído nesta lista. O Grau Principal do evento substituto e dos eventos a serem substituídos devem ser idênticos.", "E")
    q1.Active = False
    q2.Active = False
    Set q1 = Nothing
    Set q2 = Nothing
    CanContinue = False
    Exit Sub
  End If

  'Verifica se os Graus Válidos do evento a ser substituído são compatíveis com os Graus do evento substituto
  q1.Active = False
  q2.Active = False
  q1.Clear
  q2.Clear

  q1.Add("SELECT COUNT(HANDLE) NUMVALIDOS FROM SAM_TGE_GRAU WHERE EVENTO=:SUBSTITUIDO")

  q1.ParamByName("SUBSTITUIDO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  q1.Active = True

  q2.Add("SELECT COUNT(HANDLE) NUMVALIDOS")
  q2.Add("  FROM SAM_TGE_GRAU")
  q2.Add(" WHERE EVENTO = (")
  q2.Add("                 SELECT EVENTO")
  q2.Add("                   FROM SAM_TGESUBSTITUTO")
  q2.Add("                  WHERE HANDLE = :TGESUBSTITUTO")
  q2.Add("                )")
  q2.Add("   AND GRAU IN (")
  q2.Add("                SELECT GRAU")
  q2.Add("                  FROM SAM_TGE_GRAU")
  q2.Add("                 WHERE EVENTO = :EVENTO")
  q2.Add("               )")


  q2.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  q2.ParamByName("TGESUBSTITUTO").AsInteger = CurrentQuery.FieldByName("TGESUBSTITUTO").AsInteger
  q2.Active = True

  If q1.FieldByName("NUMVALIDOS").AsInteger <> q1.FieldByName("NUMVALIDOS").AsInteger Then
    bsShowMessage("Os Graus Válidos do Evento a ser substituído devem ser válidos no Evento Substituto.", "E")
    q1.Active = False
    q2.Active = False
    Set q1 = Nothing
    Set q2 = Nothing
    CanContinue = False
    Exit Sub
  End If


  'Verifica se o evento sendo incluso está inativo. Se estiver impede a inclusão.
  q1.Active = False
  q2.Active = False
  q1.Clear
  q1.Add("SELECT INATIVO FROM SAM_TGE WHERE HANDLE = :EVENTO ")
  
  q1.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  q1.Active = True

  If q1.FieldByName("INATIVO").AsString = "S" Then
    bsShowMessage("O Evento está inativo. Não é possível incluí-lo na carga de substituição.", "E")
    q1.Active = False
    Set q1 = Nothing
    CanContinue = False
    Exit Sub
  End If

  q1.Active = False
  q2.Active = False
  Set q1 = Nothing
  Set q2 = Nothing

End Sub

