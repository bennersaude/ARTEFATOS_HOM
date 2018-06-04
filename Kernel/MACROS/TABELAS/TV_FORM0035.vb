'HASH: 49B188D9891AC005C14EFCF06EBC63DE
'Macro: TV_FORM0035

'#Uses "*bsShowMessage"
Option Explicit

Public Sub EnderecoCorrespondenciaPessoa()
  Dim Query As BPesquisa
  Set Query = NewQuery

  Query.Add("SELECT ENDERECOCORRESPONDENCIA FROM SFN_PARAMETROSFIN")
  Query.Active = True

  If SessionVar("DIGITACAOENDERECO_ORIGEM") = "P" And Query.FieldByName("ENDERECOCORRESPONDENCIA").AsString = "S" Then
  	CurrentQuery.FieldByName("ENDERECO2").AsString = "S"
  End If

  Set Query = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("ENDERECO1").AsString = "S"
  CurrentQuery.FieldByName("ENDERECO2").AsString = "N"
  CurrentQuery.FieldByName("ENDERECO3").AsString = "S"
  CurrentQuery.FieldByName("ENDERECO4").AsString = "N"

  EnderecoCorrespondenciaPessoa

End Sub

Public Sub TABLE_AfterScroll()
  If SessionVar("DIGITACAOENDERECO_ORIGEM") = "B" Then
    GRUPOTIPOENDERECO.Visible = True
    ENDERECO1.Visible         = True
    ENDERECO2.Visible         = True
    ENDERECO3.Visible         = True
    ENDERECO4.Visible         = True
    ENDERECO1.Caption         = "Residencial"
    ENDERECO2.Caption         = "Comercial"
    ENDERECO3.Caption         = "Correspondência"
    ENDERECO4.Caption         = "Atend. Domiciliar"
    ENDERECO1.Hint            = "Indicar o endereço como 'Residencial'"
    ENDERECO2.Hint            = "Indicar o endereço como 'Comercial'"
    ENDERECO3.Hint            = "Indicar o endereço como 'Correspondência'"
    ENDERECO4.Hint            = "Indicar o endereço como 'Atendimento Domiciliar'"
  ElseIf SessionVar("DIGITACAOENDERECO_ORIGEM") = "P" Then
    GRUPOTIPOENDERECO.Visible = True
    ENDERECO1.Visible         = True
    ENDERECO2.Visible         = True
    ENDERECO3.Visible         = False
    ENDERECO4.Visible         = False
    ENDERECO1.Caption         = "CPF/CNPJ"
    ENDERECO2.Caption         = "Correspondência"
    ENDERECO1.Hint            = "Indicar o endereço como 'CPF/CNPJ'"
    ENDERECO2.Hint            = "Indicar o endereço como 'Correspondência'"
  ElseIf SessionVar("DIGITACAOENDERECO_ORIGEM") = "C" Then
    GRUPOTIPOENDERECO.Visible = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim lengthCPF As Integer
  Dim CEP As String
  If CurrentQuery.FieldByName("NUMERO").IsNull And _
     (CurrentQuery.FieldByName("COMPLEMENTO").IsNull Or Trim(CurrentQuery.FieldByName("COMPLEMENTO").AsString) = "") Then
   CanContinue = False
   bsShowMessage("O 'Número' ou o 'Complemento' deve estar preenchido!", "E")
  End If
  If (CurrentQuery.FieldByName("CEP").IsNull Or Trim(CurrentQuery.FieldByName("CEP").AsString) = "-") Then
   CanContinue = False
   bsShowMessage("O 'CEP' deve estar preenchido!", "E")
  End If

   CEP = Replace(CurrentQuery.FieldByName("CEP").AsString, " " , "")
   lengthCPF = Len(CEP)

  If lengthCPF <> 9 Then
   CanContinue = False
   bsShowMessage("O 'CEP'deve ser preenchido corretamente", "E")
  End If

  If SessionVar("DIGITACAOENDERECO_ORIGEM") = "B" Then
    If CurrentQuery.FieldByName("ENDERECO1").AsString <> "S" And _
       CurrentQuery.FieldByName("ENDERECO2").AsString <> "S" And _
       CurrentQuery.FieldByName("ENDERECO3").AsString <> "S" And _
       CurrentQuery.FieldByName("ENDERECO4").AsString <> "S" Then
      bsShowMessage("Deve ser marcado ao menos um tipo de endereço!", "E")
      CanContinue = False
    End If
  ElseIf SessionVar("DIGITACAOENDERECO_ORIGEM") = "P" Then
    If CurrentQuery.FieldByName("ENDERECO1").AsString <> "S" And _
       CurrentQuery.FieldByName("ENDERECO2").AsString <> "S" Then
      bsShowMessage("Deve ser marcado ao menos um tipo de endereço!", "E")
      CanContinue = False
    End If
  End If
End Sub
