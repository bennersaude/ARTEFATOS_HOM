'HASH: 99604150093D891A9D59C405C73D4233
    Public Sub BOTAOVER_OnClick()
      Dim vsArquivo As String
      Dim QWork As Object
      Dim X As Integer

      Set QWork = NewQuery

      QWork.Clear
      QWork.Add("SELECT DOCUMENTO FROM SAM_INFORMATIVO_ANEXO WHERE HANDLE = :HANDLE")
      QWork.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_INFORMATIVO_ANEXO")
      QWork.Active = True
      vsArquivo = QWork.FieldByName("DOCUMENTO").AsString
      X = RecordHandleOfTable("SAM_INFORMATIVO")
      ShowBDocFile("INFORMATIVOS\ANEXOS", HostName, vsArquivo, RecordHandleOfTable("SAM_INFORMATIVO_ANEXO"))
    End Sub


    Public Sub BOTAOSALVAR_OnClick()
      If CurrentQuery.State <>1 Then
        MsgBox("Operação Cancelada. Registro está em Edição", vbCritical)
        Exit Sub
      End If
      Dim NomeArq As String

      NomeArq = SaveDialog(CurrentQuery.FieldByName("DOCUMENTO").AsString)
      If NomeArq <>"" Then
        Dim obj As Object
        Set obj = SuperServerClient("DOC")
        obj.Select("INFORMATIVOS\ANEXOS")
        obj.GetDocument(NomeArq, CurrentQuery.FieldByName("HANDLE").AsString)
        Set obj = Nothing
      End If
    End Sub


    Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
      Dim o As Object
      On Error Resume Next
      Set o = SuperServerClient("DOC")
      o.Select("INFORMATIVOS\ANEXOS")
      o.Delete(CurrentQuery.FieldByName("HANDLE").AsString)
      o.Select("")
      Set o = Nothing
    End Sub


