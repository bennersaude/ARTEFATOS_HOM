﻿'HASH: 5CE69336111DA55BCB49E300AA1B31DF


Public Sub ATIVAR_OnClick()
  'Set obj = CreateBennerObject("ST.Ativar")
  'obj.Ambiente(CurrentSystem)RecordHandleOfTable("ST_AMBIENTEATIVIDADERISCOS")
  'obj.Codigo(CurrentSystem)2
  'obj.Exec(CurrentSystem)

  'Set obj = Nothing

End Sub

Public Sub DESATIVAR_OnClick()

  'Set obj = CreateBennerObject("ST.Desativar")
  'obj.Ambiente(CurrentSystem)RecordHandleOfTable("ST_AMBIENTEATIVIDADERISCOS")
  'obj.GrupoRisco(CurrentSystem)CurrentQuery.FieldByName("GRUPORISCO").AsInteger
  'obj.Risco(CurrentSystem)CurrentQuery.FieldByName("RISCO").AsInteger
  'obj.Codigo(CurrentSystem)2
  'obj.Exec(CurrentSystem)

  'Set obj = Nothing

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Sql As Object
  Dim Tipo As Integer
  Dim Nivelaceitavel As String, nivelacao As String, nivelmonitoramento As String
  Dim Valor, Metade, Valor2, ValorDigitado

  If Not(CurrentQuery.FieldByName("VALORENCONTRADO").IsNull)Then
    Set Sql = NewQuery

    Sql.Add "SELECT  TIPORISCO, LIMITE, LIMITEACAO, LIMITEMONITORAMENTO FROM ST_RISCOS WHERE  HANDLE = " + CurrentQuery.FieldByName("RISCO").AsString
    Sql.Active = True
    Tipo = Sql.FieldByName("TIPORISCO").Value
    If(Tipo = 2)Or(Tipo = 3)Then
    Valor = Sql.FieldByName("LIMITE").Value
    Nivelaceitavel = "Nível aceitável"
    nivelacao = "Nível de ação (> LT)"
    nivelmonitoramento = "Nível de ação (<LT)"
    ValorDigitado = CurrentQuery.FieldByName("VALORENCONTRADO").Value
    If Not Sql.EOF And Tipo = 2 Then
      If Valor >0 Then
        Metade = Valor / 2
        If ValorDigitado <Metade Then
          CurrentQuery.FieldByName("SITUACAO").Value = Nivelaceitavel
          CurrentQuery.FieldByName("TIPOSITUACAO").Value = 3
        Else
          If((ValorDigitado >Metade Or ValorDigitado = Metade)And ValorDigitado <Valor)Then
          CurrentQuery.FieldByName("SITUACAO").Value = nivelmonitoramento
          CurrentQuery.FieldByName("TIPOSITUACAO").Value = 2
        Else
          If(ValorDigitado >Valor Or ValorDigitado = Valor)Then
          CurrentQuery.FieldByName("SITUACAO").Value = nivelacao
          CurrentQuery.FieldByName("TIPOSITUACAO").Value = 1
        End If
      End If
    End If
  Else
    MsgBox "Risco sem valor do limite de tolerância"
    CanContinue = False
  End If
Else
  Valor = Sql.FieldByName("LIMITEACAO").Value
  Valor2 = Sql.FieldByName("LIMITEMONITORAMENTO").Value
  If Not Sql.EOF And Tipo = 3 Then
    If((ValorDigitado >Valor2)Or(ValorDigitado = Valor2))And(ValorDigitado <Valor)Then
    CurrentQuery.FieldByName("SITUACAO").Value = nivelmonitoramento
    CurrentQuery.FieldByName("TIPOSITUACAO").Value = 2
  Else
    If((ValorDigitado >Valor)Or(ValorDigitado = Valor))Then
    CurrentQuery.FieldByName("SITUACAO").Value = nivelacao
    CurrentQuery.FieldByName("TIPOSITUACAO").Value = 1
  Else
    CurrentQuery.FieldByName("SITUACAO").Value = " "
  End If
End If
End If
End If
Else
  MsgBox "Risco qualitativo"
  CanContinue = False
End If
End If
End Sub
