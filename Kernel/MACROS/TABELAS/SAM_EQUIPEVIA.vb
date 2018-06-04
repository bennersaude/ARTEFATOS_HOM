'HASH: C68B569A37CFD7B75AFBF210507EF8B9
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("BILATERALANESTESISTA").AsInteger = 1 And _
                               CurrentQuery.FieldByName("PERCENTUALANEST").IsNull Then
    bsShowMessage("Campo 'Cód. pagto equipe anestesista' obrigatório quando o " + Chr(13) + _
      "'Perc. difer. anestes. evento bilateral' for igual a SIM !", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

