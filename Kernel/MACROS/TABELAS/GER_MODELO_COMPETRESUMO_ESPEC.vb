'HASH: 98EF4C50B3406ECCBC500920EB39DF48
 

Public Sub TABLE_AfterScroll()
  ESPECIALIDADE.Visible = False
  TIPOTRATAMENTO.Visible = False

  If  CurrentQuery.FieldByName("ESPECIALIDADE").IsNull Then
    TIPOTRATAMENTO.Visible = True
  End If

  If CurrentQuery.FieldByName("TIPOTRATAMENTO").IsNull Then
    ESPECIALIDADE.Visible = True
  End If

End Sub
