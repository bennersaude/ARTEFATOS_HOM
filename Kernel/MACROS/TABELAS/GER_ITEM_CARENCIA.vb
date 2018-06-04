'HASH: 34DC9B70BFDD40174727F0FDE6347A94
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim q As Object
  Set q = NewQuery
  q.Clear

  q.Add("SELECT 1 FROM ger_item_carencia WHERE grupocarencia=:grupocarencia AND item=:item ")

  q.ParamByName("grupocarencia").Value = CurrentQuery.FieldByName("grupocarencia").Value
  q.ParamByName("item"         ).Value = CurrentQuery.FieldByName("item"         ).Value
  q.Active = True

  If Not q.EOF Then
    CanContinue = False
   	MsgBox("Este grupo de carência já está cadastrado para este item!")

  End If

End Sub
