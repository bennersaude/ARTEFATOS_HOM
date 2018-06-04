'HASH: EA98667746B377377E33D665ED2AE667
Public Function SuportMaiusculaMinuscula As Boolean 
	SuportMaiusculaMinuscula = NewMemoryTableByName("Z_GRUPOUSUARIOS").FieldByName("SENHA").Largura >= 100 
End Function 
 
Public Sub TABLE_AfterScroll() 
	MAIUSCULASSENHA.Visible = SuportMaiusculaMinuscula 
	MINUSCULASSENHA.Visible = SuportMaiusculaMinuscula 
End Sub 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
 
  ' Validação da complexidade da senha 
 
  Dim TamanhoMin As Long 
  Dim TamanhoMax As Long 
 
  TamanhoMin = CurrentQuery.FieldByName("TAMANHOSENHA").AsInteger 
  TamanhoMax = CurrentQuery.FieldByName("TAMANHOSENHAMAX").AsInteger 
 
  MinimoEsperadoSenha = CurrentQuery.FieldByName("NUMEROSSENHA").AsInteger + CurrentQuery.FieldByName("ESPECIAISSENHA").AsInteger 
  MinimoMaiusculaMinusculaSenha = CurrentQuery.FieldByName("MAIUSCULASSENHA").AsInteger + CurrentQuery.FieldByName("MINUSCULASSENHA").AsInteger 
 
  If MinimoMaiusculaMinusculaSenha > CurrentQuery.FieldByName("LETRASSENHA").AsInteger Then 
	MinimoEsperadoSenha = MinimoEsperadoSenha + MinimoMaiusculaMinusculaSenha 
  Else 
    MinimoEsperadoSenha = MinimoEsperadoSenha + CurrentQuery.FieldByName("LETRASSENHA").AsInteger 
  End If 
 
 
  If Not CurrentQuery.FieldByName("TAMANHOSENHA").IsNull Then 
	If Not CurrentQuery.FieldByName("TAMANHOSENHAMAX").IsNull Then 
	  If TamanhoMax < TamanhoMin Then 
        CanContinue = False 
        CancelDescription = "Tamanho máximo da senha deve ser maior ou igual ao tamanho mínimo" 
        If VisibleMode Then 
          MsgBox(CancelDescription) 
        End If 
      	Exit Sub 
      End If 
    End If 
 
    If (TamanhoMin < MinimoEsperadoSenha) Then 
      CanContinue = False 
      If MinimoMaiusculaMinusculaSenha > 0 Then 
	    CancelDescription = "O tamanho mínimo não pode ser menor que a soma do mínimo de números, letras, maiusculas e minusculas. (" & MinimoEsperadoSenha & ")" 
      Else 
		CancelDescription = "O tamanho mínimo não pode ser menor que a soma do mínimo de números, letras. (" & MinimoEsperadoSenha & ")" 
      End If 
      If VisibleMode Then 
        MsgBox(CancelDescription) 
      End If 
      Exit Sub 
    End If 
 
  End If 
 
  If Not CurrentQuery.FieldByName("TAMANHOSENHAMAX").IsNull Then 
 
    If TamanhoMax < TamanhoMin Then 
      CanContinue = False 
      CancelDescription = "Tamanho máximo da senha deve ser maior ou igual ao tamanho mínimo" 
      If VisibleMode Then 
        MsgBox(CancelDescription) 
      End If 
      Exit Sub 
    End If 
 
    If (TamanhoMax < MinimoEsperadoSenha) Then 
      CanContinue = False 
      If MinimoMaiusculaMinusculaSenha > 0 Then 
	    CancelDescription = "O tamanho mínimo não pode ser menor que a soma do mínimo de números, letras, maiusculas e minusculas. (" & MinimoEsperadoSenha & ")" 
      Else 
		CancelDescription = "O tamanho mínimo não pode ser menor que a soma do mínimo de números, letras. (" & MinimoEsperadoSenha & ")" 
      End If 
      If VisibleMode Then 
        MsgBox(CancelDescription) 
      End If 
      Exit Sub 
    End If 
 
  End If 
 
 
  ' Validação do histórico de senha 
 
  If (CurrentQuery.FieldByName("HISTORICO").AsInteger < 2) Then 
    CurrentQuery.FieldByName("NAOREPETIRSENHA").Clear 
  ElseIf (CurrentQuery.FieldByName("HISTORICO").AsInteger = 2) Then 
    CurrentQuery.FieldByName("NAOREPETIRSENHA").AsInteger = 0 
  ElseIf (CurrentQuery.FieldByName("HISTORICO").AsInteger = 3) Then 
    If (CurrentQuery.FieldByName("NAOREPETIRSENHA").AsInteger < 1) Then 
      CanContinue = False 
      CancelDescription = "A quantidade de senhas deve ser maior que 0" 
      If VisibleMode Then 
        MsgBox(CancelDescription) 
      End If 
      Exit Sub 
    End If 
  End If 
 
End Sub 
