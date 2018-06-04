'HASH: 6FBFBFAE6F3E5F4C8F288CA6212C5B77
'#Uses "*bsShowMessage"
'Daniela Zardo - 17/07/2002

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Or 2 Then
    If CurrentQuery.FieldByName("OBRIGATORIO").AsString = "S" Then
      Dim q1 As Object
      Set q1 = NewQuery

      q1.Add("SELECT HANDLE FROM SAM_PLANO WHERE HANDLE = :HPLANO AND DECOMPOSICAOFAMILIAR = 'S' ")
      q1.ParamByName("HPLANO").Value = RecordHandleOfTable("SAM_PLANO")
      q1.Active = True
      If q1.EOF Then
        bsShowMessage("Esse plano não permite decomposição familiar", "E")
        CanContinue = False
      End If
    End If
  End If

End Sub

