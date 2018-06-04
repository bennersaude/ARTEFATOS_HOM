'HASH: 1FA50105080CFB4C4AF1C00C1404B580
  '#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not VisibleMode Then
    Exit Sub
  End If

  Dim qIntervencao As Object
  Set qIntervencao = NewQuery

  If CurrentQuery.FieldByName("PERMITEGRUPO").AsString = "S" Then

    qIntervencao.Active = False

    qIntervencao.Clear
    qIntervencao.Add("SELECT HANDLE ")
    qIntervencao.Add("  FROM CLI_INTERVENCAO ")
    qIntervencao.Add(" WHERE HANDLE = :INTERVENCAO ")
    qIntervencao.Add("   AND HANDLE IN (3, 6, 12, 18, 19) ")
    qIntervencao.ParamByName("INTERVENCAO").AsInteger = CurrentQuery.FieldByName("INTERVENCAO").AsInteger

    qIntervencao.Active = True

    If qIntervencao.EOF Then
      bsShowMessage("O campo 'Permite informar para grupo' somente pode ser marcado para as intervenções de Raio X, Periodontia e Ponte.", "E")
      CanContinue = False
    End If

  Else
    qIntervencao.Active = False

    qIntervencao.Clear
    qIntervencao.Add("SELECT HANDLE ")
    qIntervencao.Add("  FROM CLI_INTERVENCAO ")
    qIntervencao.Add(" WHERE HANDLE = :INTERVENCAO ")
    qIntervencao.Add("   AND HANDLE IN (12, 18, 19) ")
    qIntervencao.ParamByName("INTERVENCAO").AsInteger = CurrentQuery.FieldByName("INTERVENCAO").AsInteger

    qIntervencao.Active = True

    If Not qIntervencao.EOF Then
      CurrentQuery.FieldByName("PERMITEGRUPO").AsString = "S"
    End If
  End If

  Set qIntervencao = Nothing
End Sub

