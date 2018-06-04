'HASH: BE388778A186990FBA53FEF7E276DD53

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim q As Object
  Set q = NewQuery
  q.Clear

  q.Add("SELECT 1 FROM ger_item_camposip WHERE camposip=:camposip AND grupobenef=:grupobenef AND item=:item ")

  q.ParamByName("camposip"  ).Value = CurrentQuery.FieldByName("camposip"  ).Value
  q.ParamByName("grupobenef").Value = CurrentQuery.FieldByName("grupobenef").Value
  q.ParamByName("item"      ).Value = CurrentQuery.FieldByName("item"      ).Value
  q.Active = True

  If Not q.EOF Then
    CanContinue = False
   	MsgBox("Campo sip já cadastrado para esse item !")
  End If
  Set q = Nothing

End Sub
