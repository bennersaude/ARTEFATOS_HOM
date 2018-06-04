'HASH: 3793FE960BA8CFE23B7240987B435FFC
'CLI_PARTICIPANTE
Option Explicit
Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vWhere As String
  Dim vColunas As String

  Dim vCabecalho As String
  ShowPopup = False

  vCabecalho = "Nome|Beneficiario|Código Afinidade|Código Antigo|Data Cancelamento"
  vColunas = "SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_BENEFICIARIO.CODIGODEAFINIDADE|SAM_BENEFICIARIO.CODIGOANTIGO|SAM_BENEFICIARIO.DATACANCELAMENTO"
  vWhere = vWhere + "(ATENDIMENTOATE IS NULL OR ATENDIMENTOATE >= " + SQLDate(ServerNow) + ") AND "
  vWhere = vWhere + "(DATABLOQUEIO IS NULL) And "
  vWhere = vWhere + "(DATACANCELAMENTO IS NULL OR DATACANCELAMENTO >= " + SQLDate(ServerNow) + ") "
  Set interface = CreateBennerObject("Procura.Procurar")
  CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 1, vCabecalho, vWhere, "Procura por beneficiário", False, "")
  Set interface = Nothing
End Sub

Public Sub MATRICULA_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vWhere As String
  Dim vColunas As String
  Dim vData As String
  Dim vCabecalho As String
  ShowPopup = False
  vCabecalho = "Nome|CPF|Data de nascimento|RG|Data de ingresso"
  vColunas = "SAM_MATRICULA.Z_NOME|SAM_MATRICULA.CPF|SAM_MATRICULA.DATANASCIMENTO|SAM_MATRICULA.RG|SAM_MATRICULA.DATAINGRESSO"
  vWhere = ""
  Set interface = CreateBennerObject("Procura.Procurar")
  CurrentQuery.FieldByName("MATRICULA").AsInteger = interface.Exec(CurrentSystem, "SAM_MATRICULA", vColunas, 1, vCabecalho, vWhere, "Procura por matrícula", False, "")
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Select Case NodeInternalCode
    Case 2640
      CurrentQuery.FieldByName("TABTIPOPARTICIPANTE").AsInteger = 1
    Case 2650
    CurrentQuery.FieldByName("TABTIPOPARTICIPANTE").AsInteger = 2
  End Select
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim QUANTIDADE As Integer
    Dim QTD As Object
    Set QTD = NewQuery
    QTD.Clear
    QTD.Add("SELECT QTDPARTICIPANTES FROM CLI_TURMA WHERE HANDLE = :TURMA")
    QTD.ParamByName("TURMA").Value = CurrentQuery.FieldByName("TURMA").AsInteger
    QTD.Active = True
    QUANTIDADE = QTD.FieldByName("QTDPARTICIPANTES").AsInteger

    If QUANTIDADE >0 Then
      Dim TOTAL As Integer
      Dim VERIFICA As Object
      Set VERIFICA = NewQuery

      VERIFICA.Clear
      VERIFICA.Add("SELECT COUNT(HANDLE) TOTAL FROM CLI_TURMA_PARTICIPANTE WHERE TURMA = :TURMA")
      VERIFICA.ParamByName("TURMA").Value = CurrentQuery.FieldByName("TURMA").AsInteger
      VERIFICA.Active = True
      TOTAL = VERIFICA.FieldByName("TOTAL").AsInteger

      If TOTAL >= QUANTIDADE Then
        MsgBox("A quantidade máxima de participantes desta turma não pode exceder " + _
               Str(QUANTIDADE) + " participantes!")
        CanContinue = False
        Exit Sub
      End If
      Set VERIFICA = Nothing
    End If
    Set QTD = Nothing
  End If

  If (CurrentQuery.FieldByName("TABTIPOPARTICIPANTE").AsInteger = 1 And (CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = 0 Or CurrentQuery.FieldByName("BENEFICIARIO").AsString = "")) Then
  	MsgBox("Campo 'Beneficiário' é obrigatório!")
    CanContinue = False
    Exit Sub
  ElseIf (CurrentQuery.FieldByName("TABTIPOPARTICIPANTE").AsInteger = 2 And	(CurrentQuery.FieldByName("MATRICULA").AsInteger = 0 Or CurrentQuery.FieldByName("MATRICULA").AsString = "")) Then
	MsgBox("Campo 'Matrícula' é obrigatório!")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("TABTIPOPARTICIPANTE").AsInteger = 1) Then
    CurrentQuery.FieldByName("MATRICULA").Clear
  ElseIf (CurrentQuery.FieldByName("TABTIPOPARTICIPANTE").AsInteger = 2) Then
    CurrentQuery.FieldByName("BENEFICIARIO").Clear
  End If

End Sub

