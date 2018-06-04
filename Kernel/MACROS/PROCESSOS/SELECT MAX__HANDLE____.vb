'HASH: 9F21263917A690CAB03B5A6C8B4E4F3C
Sub Main
  Dim Sql As Object
  Set Sql = NewQuery

  Dim Del As Object
  Set Del = NewQuery

  Sql.Add("SELECT NOME ")
  Sql.Add("  FROM Z_TABELAS")
  Sql.Add("  ORDER BY NOME ")
  Sql.Active = True
  
  Dim mail as object
  Set mail = NewMail

    mail.subject="tabelas da base PRO"
    mail.From="ricardo@bennersaude.com.br"
    mail.SendTo="ricardo@bennersaude.com.br"

  While Not Sql.EOF
    Del.clear
    Del.Add("SELECT MAX(HANDLE) handlemaximo FROM  "+ Sql.FieldByName("NOME").AsString)
    Del.ACTIVE = TRUE
    
    If Del.FieldByName("HANDLEMAXIMO").AsInteger > 0 Then
        mail.Text.Add(Chr(13) + sql.fieldbyname("NOME").asstring + " ; " + Del.FieldByName("HANDLEMAXIMO").asstring)

    Sql.Next
  Wend

     mail.Send

  Set Sql = Nothing
  Set Del = Nothing
  set mail = Nothing

End sub
