'HASH: 3ED3F5059243386032FAF3158DAD864E
'Macro: SAM_MATRICULALOTE
'#Uses "*bsShowMessage"

Public Sub BOTAOIMPORTAR_OnClick()
  Dim Importar As Object

  If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
    bsShowMessage("Operação cancelada. Usuário não é o Responsável", "I")
    Exit Sub
  End If

  If CurrentQuery.State = 1 And CurrentQuery.FieldByName("LOTEPROCESSADO").AsString = "N" Then 'Não está em inserção
    Set Importar = CreateBennerObject("Matricula.Importacao")
    Importar.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Importa = Nothing

    WriteAudit("I", HandleOfTable("SAM_MATRICULALOTE"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Matrículas Provisórias - Importação")
  Else
    bsShowMessage("O lote nâo pode estar em inserção, ou lote já processado", "I")
  End If

End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Gravar As Object

  If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
    bsShowMessage("Operação cancelada. Usuário não é o Responsável", "I")
    Exit Sub
  End If

  If CurrentQuery.State = 1 And CurrentQuery.FieldByName("LOTEPROCESSADO").AsString = "N" Then 'Não está em inserção
    Set Gravar = CreateBennerObject("Matricula.Importacao")
    Gravar.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    RefreshNodesWithTable("SAM_MATRICULALOTE")
    Set Gravar = Nothing

    WriteAudit("P", HandleOfTable("SAM_MATRICULALOTE"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Matrículas Provisórias - Processamento")
  Else
    bsShowMessage("O lote nâo pode estar em inserção, ou lote já processado", "I")
  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim vFrase As String
  Dim SQL As Object
  Dim SQLMAT As Object

  If CurrentUser = CurrentQuery.FieldByName("USUARIO").AsInteger Then

    Set DEL = NewQuery


    DEL.Clear
    DEL.Add("DELETE FROM SAM_MATRICULAHOMONIMA ")
    DEL.Add(" WHERE MATRICULAPROVISORIA IN (SELECT HANDLE ")
    DEL.Add("                                 FROM SAM_MATRICULAPROVISORIA ")
    DEL.Add("                                WHERE LOTE = :LOTE)")
    DEL.ParamByName("LOTE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    DEL.ExecSQL

    DEL.Clear
    DEL.Add("DELETE FROM SAM_MATRICULAPROVISORIA WHERE LOTE = :LOTE")
    DEL.ParamByName("LOTE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    DEL.ExecSQL

    Set DEL = Nothing



  Else

    CanContinue = False
    bsShowMessage("Operação cancelada. Usuário não é o Responsável", "E")

  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
    CanContinue = False
    bsShowMessage("Operação cancelada. Usuário não é o Responsável", "E")
  End If

End Sub



Public Sub TABLE_NewRecord()
  Dim prFilial As Long
  Dim prFilialProcessamento As Long
  Dim prMsg As String
  BuscarFiliais(CurrentSystem, prFilial, prFilialProcessamento, prMsg)
  CurrentQuery.FieldByName("filial").AsInteger = prFilial

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOIMPORTAR"
			BOTAOIMPORTAR_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
