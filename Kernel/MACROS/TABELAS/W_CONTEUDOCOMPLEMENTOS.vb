'HASH: 3EE744B50F0CB6C25EFA6A8754F4739C
Public Sub TABLE_AfterScroll() 
  CODIGOINTERNO.Text = "Código interno: " + CurrentQuery.FieldByName("HANDLE").AsString 
End Sub 
