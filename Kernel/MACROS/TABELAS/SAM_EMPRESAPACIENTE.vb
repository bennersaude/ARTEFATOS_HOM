'HASH: 3B34354080DC15603280B4CD057E4C39

Dim vsCodigo As String

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vsCodigo = UCase(CurrentQuery.FieldByName("CODIGO").Value)
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim SQL1 As Object
  Dim SQL2 As Object
  Set SQL = NewQuery
  Set SQL1 = NewQuery
  Set SQL2 = NewQuery


  If CurrentQuery.State = 3 Then 'INSERT
    SQL.Clear
    SQL.Add (" SELECT CODIGO FROM SAM_EMPRESAPACIENTE WHERE CODIGO = :CODIGO ")
    SQL.ParamByName("CODIGO").Value = UCase(CurrentQuery.FieldByName("CODIGO").Value)
    SQL.Active = True

    If Not (SQL.FieldByName("CODIGO").IsNull) Then
      MsgBox("Já existe este código. Favor verificar ")
      CanContinue = False
    End If
  End If

  If CurrentQuery.State = 2 Then 'EDIT

    If vsCodigo <> UCase(CurrentQuery.FieldByName("CODIGO").Value) Then
      SQL.Clear
      SQL.Add (" SELECT CODIGO FROM SAM_EMPRESAPACIENTE WHERE CODIGO = :CODIGO ")
      SQL.ParamByName("CODIGO").Value = UCase(CurrentQuery.FieldByName("CODIGO").Value)
      SQL.Active = True

      If Not (SQL.FieldByName("CODIGO").IsNull) Then
        MsgBox("Já existe este código. Favor verificar ")
        CanContinue = False
      End If
    End If
  End If


  If CurrentQuery.FieldByName("SITUACAO").AsString = "I" Then

    SQL1.Clear
    SQL1.Add (" SELECT *                                                  ")
    SQL1.Add ("   FROM SAM_MATRICULA_EMPRESAPACIENTE                      ")
    SQL1.Add ("  WHERE DATAINICIAL <= :DATA                               ")
    SQL1.Add ("    AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)          ")
    SQL1.Add ("    AND EMPRESAPACIENTE = :EMPRESA                         ")
    SQL1.ParamByName("DATA").Value = ServerDate
    SQL1.ParamByName("EMPRESA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL1.Active = True

    If Not (SQL1.FieldByName("HANDLE").IsNull) Then

      If MsgBox("Esta empresa está ligada a pelo menos uma matrícula com vigência válida. Deseja alterar sua situação para 'Inativa'? ", vbQuestion + vbYesNo, "") = vbYes Then

        While Not (SQL1.EOF)
          SQL2.Clear
          SQL2.Add (" UPDATE SAM_MATRICULA_EMPRESAPACIENTE SET DATAFINAL = :DATA WHERE EMPRESAPACIENTE = :EMPRESA ")
          SQL2.ParamByName("DATA").Value = ServerDate
          SQL2.ParamByName("EMPRESA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
          SQL2.ExecSQL

          SQL1.Next
        Wend

      Else
        CurrentQuery.FieldByName("SITUACAO").AsString = "A"
      End If
    End If
  End If

  CurrentQuery.FieldByName("CODIGO").Value = UCase(CurrentQuery.FieldByName("CODIGO").Value)

End Sub

