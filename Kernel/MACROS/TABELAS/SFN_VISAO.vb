'HASH: 78B37E561BFE9911D8F0B8CD9D80652B
'Macro: SFN_VISAO
Option Explicit


Public Sub BOTAOCOPIAR_OnClick()
  Dim OldArray()As Long 'Guarda o o velho handle
  Dim NewArray()As Long 'guarda o novo handle
  Dim Q As Object
  Dim S As Object
  Dim VHandle As Long
  Dim EHandle As Long
  Dim Quant As Long
  Dim I As Long
  Dim X As Long

  Set Q = NewQuery
  Set S = NewQuery

  If Not InTransaction Then StartTransaction

  ' Cria a visao
  VHandle = NewHandle("SFN_VISAO")
  Q.Clear
  Q.Add("INSERT INTO SFN_VISAO (HANDLE,DESCRICAO) VALUES (:HANDLE,:DESCRICAO)")
  Q.ParamByName("HANDLE").Value = VHandle
  Q.ParamByName("DESCRICAO").Value = CurrentQuery.FieldByName("DESCRICAO").AsString + " " + Str(VHandle)
  Q.ExecSQL

  ' Conta os registros
  S.Clear
  S.Add("SELECT COUNT(*) QUANTIDADE FROM SFN_VISAOESTRUTURA WHERE VISAO=:VISAO")
  S.ParamByName("VISAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  S.Active = True
  Quant = S.FieldByName("QUANTIDADE").AsInteger
  S.Active = False

  ' Monta uma query
  S.Clear
  S.Add("SELECT HANDLE,ESTRUTURA,NATUREZA,DESCRICAO,NIVELSUPERIOR,ULTIMONIVEL FROM SFN_VISAOESTRUTURA WHERE VISAO=:VISAO ORDER BY VISAO,ESTRUTURA")
  S.ParamByName("VISAO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  S.Active = True

  ReDim OldArray(Quant)
  ReDim NewArray(Quant)
  I = 0
  While Not S.EOF
    EHandle = NewHandle("SFN_VISAOESTRUTURA")
    OldArray(I) = S.FieldByName("HANDLE").AsInteger
    NewArray(I) = EHandle
    I = I + 1

    Q.Clear
    If S.FieldByName("NIVELSUPERIOR").AsInteger = 0 Then
      Q.Add("INSERT INTO SFN_VISAOESTRUTURA (HANDLE,VISAO,ESTRUTURA,NATUREZA,DESCRICAO,NIVELSUPERIOR,ULTIMONIVEL) VALUES (:HANDLE,:VISAO,:ESTRUTURA,:NATUREZA,:DESCRICAO,NULL,:ULTIMONIVEL)")
    Else
      Q.Add("INSERT INTO SFN_VISAOESTRUTURA (HANDLE,VISAO,ESTRUTURA,NATUREZA,DESCRICAO,NIVELSUPERIOR,ULTIMONIVEL) VALUES (:HANDLE,:VISAO,:ESTRUTURA,:NATUREZA,:DESCRICAO,:NIVELSUPERIOR,:ULTIMONIVEL)")
    End If
    Q.ParamByName("HANDLE").Value = EHandle
    Q.ParamByName("VISAO").Value = VHandle
    Q.ParamByName("ESTRUTURA").Value = S.FieldByName("ESTRUTURA").AsString
    Q.ParamByName("NATUREZA").Value = S.FieldByName("NATUREZA").AsString
    Q.ParamByName("DESCRICAO").Value = S.FieldByName("DESCRICAO").AsString
    Q.ParamByName("ULTIMONIVEL").Value = S.FieldByName("ULTIMONIVEL").AsString

    If S.FieldByName("NIVELSUPERIOR").AsInteger <>0 Then
      X = 0
      While X <= Quant
        If S.FieldByName("NIVELSUPERIOR").AsInteger = OldArray(X)Then
          Q.ParamByName("NIVELSUPERIOR").AsInteger = NewArray(X)
          X = Quant 'Para forçar a saida do loop
        End If
        X = X + 1
      Wend
    End If

    Q.ExecSQL

    S.Next

  Wend

  If InTransaction Then Commit

  RefreshNodesWithTable("SFN_VISAO")

End Sub

Public Sub BOTAOCONFIGURAR_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SfnGerencial.Rotinas")
  interface.GeradorVisao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
      Case "BOTAOCOPIAR"
        BOTAOCOPIAR_OnClick
  End Select
End Sub
