'HASH: 142DC8F8E199695D07E4A7F0F44FB558
'SAM_GUIA_DEVOLUCAO_GUIA

Public Sub TABLE_AfterScroll()
  Dim q1 As Object
  Dim q2 As Object
  Set q1 = NewQuery
  Set q2 = NewQuery
  q2.Add("SELECT * FROM SAM_GUIA_DEVOLUCAO WHERE HANDLE=:HANDLE")
  q2.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("GUIADEVOLUCAO").AsInteger
  q2.Active = True

  If q2.FieldByName("TABDEVOLUCAO").AsInteger = 1 Then
    q1.Add("SELECT P.NOME FROM SAM_PRESTADOR P WHERE P.HANDLE=:HANDLE")
    q1.ParamByName("HANDLE").Value = q2.FieldByName("PRESTADOR").AsInteger
    aux = "Recebedor: "
  ElseIf q2.FieldByName("TABDEVOLUCAO").AsInteger = 2 Then
    q1.Add("SELECT P.NOME FROM SAM_BENEFICIARIO P WHERE P.HANDLE=:HANDLE")
    q1.ParamByName("HANDLE").Value = q2.FieldByName("BENEFICIARIO").AsInteger
    aux = "Beneficiário: "
  Else
    Set q1 = Nothing
    Set q2 = Nothing
    RECEBEDOR.Text = ""
    Exit Sub
  End If

  q1.Active = True
  RECEBEDOR.Text = aux + q1.FieldByName("NOME").AsString
  Set q1 = Nothing
  Set q2 = Nothing

' toda a tabela deve ser somente para leitura, nenhum campo pode ser editavel
  GUIA.ReadOnly = True
  COMPROVANTEATENDIMENTO.ReadOnly = True
  ACEITAREGULARIZACAO.ReadOnly = True
  GUIAPROCESSADA.ReadOnly = True
  GUIADEVOLUCAO.ReadOnly = True

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery
  Dim q2 As Object
  Set q2 = NewQuery
  'Marcelo Barbosa - SMS 29392 - 04/10/2005

  If (Not CurrentQuery.FieldByName("GUIA").IsNull)	And (Not CurrentQuery.FieldByName("COMPROVANTEATENDIMENTO").IsNull) Then
    MsgBox("Não é permitido informar a guia e o comprovante de atendimento ao mesmo tempo.",vbOkOnly)
    CanContinue = False
  ElseIf CurrentQuery.FieldByName("GUIA").IsNull And CurrentQuery.FieldByName("COMPROVANTEATENDIMENTO").IsNull Then
     MsgBox("Deve informar o número da guia ou o comprovante de atendimento.",vbOkOnly)
     CanContinue = False
  Else
  	q2.Clear
  	q2.Add("SELECT PEG FROM SAM_GUIA_DEVOLUCAO WHERE HANDLE=:GUIADEVOLUCAO")
  	q2.ParamByName("GUIADEVOLUCAO").Value = CurrentQuery.FieldByName("GUIADEVOLUCAO").AsInteger
  	q2.Active = True

    If Not CurrentQuery.FieldByName("GUIA").IsNull Then
   	  q1.Clear
  	  q1.Add("SELECT GUIA FROM SAM_GUIA WHERE PEG=:PEG AND GUIA=:GUIA")
  	  q1.ParamByName("PEG").Value = q2.FieldByName("PEG").AsInteger
  	  q1.ParamByName("GUIA").Value = CurrentQuery.FieldByName("GUIA").AsFloat
  	  q1.Active = True
  	Else
  	  q1.Clear
  	  q1.Add("SELECT COMPROVANTEATENDIMENTO FROM SAM_GUIA WHERE PEG=:PEG AND COMPROVANTEATENDIMENTO=:COMPROVANTEATENDIMENTO")
  	  q1.ParamByName("PEG").Value = q2.FieldByName("PEG").AsInteger
  	  q1.ParamByName("COMPROVANTEATENDIMENTO").Value = CurrentQuery.FieldByName("COMPROVANTEATENDIMENTO").AsString
  	  q1.Active = True
  	End If

  	If Not q1.EOF Then
      If MsgBox("Existe uma guia com este número para o respectivo PEG. Continuar mesmo assim?", 4) = vbYes Then
        CanContinue = True
      Else
        CanContinue = False
      End If
    Else
      CanContinue = True
    End If
  End If
  'fim - SMS 29392

  q1.Active = False
  q2.Active = False
  Set q1 = Nothing
  Set q2 = Nothing

End Sub

Public Sub TABLE_NewRecord()
  Dim PRFILIAL As Long
  Dim PRFILIALPROCESSAMENTO As Long
  Dim PRMSG As String
  Dim q2 As Object
  Set q2 = NewQuery

  If BuscarFiliais(CurrentSystem, PRFILIAL, PRFILIALPROCESSAMENTO, PRMSG) Then
    MsgBox ("Erro na rotina de Busca de filiais.")
    Exit Sub
  End If

  CurrentQuery.FieldByName("filialprocessamento").Value = PRFILIAL

  q2.Clear
  q2.Add("SELECT PEG.PEG                                                   ")
  q2.Add("  FROM SAM_PEG PEG                                               ")
  q2.Add("  JOIN SAM_GUIA_DEVOLUCAO      P ON (P.PEG = PEG.HANDLE)         ")
  q2.Add("  JOIN SAM_GUIA_DEVOLUCAO_GUIA G ON (G.GUIADEVOLUCAO = P.HANDLE) ")
  q2.Add(" WHERE P.HANDLE = :GUIADEVOLUCAO                                 ")
  q2.ParamByName("GUIADEVOLUCAO").Value = RecordHandleOfTable("SAM_GUIA_DEVOLUCAO")
  q2.Active = True

  CurrentQuery.FieldByName("PEGORIGEM").Value = q2.FieldByName("PEG").AsInteger

End Sub

