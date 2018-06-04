'HASH: 2B87DA0BF4F8D6666AE3CCDC541775F1
'Macro tabela: MS_MEDICOCOORDENADOR

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim vColuna As String
  Dim vCriterio As String
  Dim vCampos As String
  Dim vHandle As Integer
  Dim interface As Object

  Set interface = CreateBennerObject("Procura.Procurar")

  ShowPopup = False

  vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
  vCriterio = "HANDLE IN (SELECT DISTINCT PRESTADOR FROM CLI_RECURSO)"
  vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Médico", False, PRESTADOR.Text)

  If vHandle > 0 Then CurrentQuery.FieldByName("PRESTADOR").AsInteger = vHandle

  Set interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime) Then
      MsgBox "A data inicial não pode ser superior à data final!"
      CanContinue = False
      Exit Sub
    End If
  End If

  Dim Verifica As Object

  Set Verifica = NewQuery

  Verifica.Active = False
  Verifica.Clear
  Verifica.Add("SELECT HANDLE                          ")
  Verifica.Add("  FROM MS_MEDICOCOORDENADOR            ")
  Verifica.Add(" WHERE HANDLE <> :HANDLE               ")

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    Verifica.Add("   AND (   (DATAINICIAL <= :DATAINI AND DATAFINAL >= :DATAFIM)")
    Verifica.Add("        OR (DATAINICIAL >  :DATAINI AND ((DATAINICIAL <= :DATAFIM AND DATAFINAL >= :DATAFIM)))")
    Verifica.Add("        OR (DATAFINAL <  :DATAFIM AND ((DATAINICIAL <= :DATAINI AND DATAFINAL >= :DATAINI)))")
    Verifica.Add("        OR (DATAINICIAL > :DATAINI AND DATAFINAL < :DATAFIM)")
    Verifica.Add("        OR (DATAINICIAL <= :DATAFIM AND DATAFINAL IS NULL))")
    Verifica.ParamByName("DATAFIM").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  Else
    Verifica.Add("   AND (   (DATAFINAL IS NULL)")
    Verifica.Add("        OR (DATAFINAL >= :DATAINI))")
  End If

  Verifica.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  Verifica.ParamByName("DATAINI").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  Verifica.Active = True

  If Not Verifica.EOF Then
    MsgBox "Já existe um médico coordenador cadastrado para esta vigência!"
    Set Verifica = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set Verifica = Nothing

End Sub
 
