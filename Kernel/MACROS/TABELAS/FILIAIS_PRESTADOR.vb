'HASH: 702C42A0F8FA8CCF38FAF87C2DFE0ABB
'MACRO: FILIAIS_PRESTADOR


Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Dim vOrdemBusca As Integer
  Dim vbAux As Boolean


  vbAux = True
  On Error GoTo caracteres
  CDbl(PRESTADOR.LocateText)
  vOrdemBusca = 1
  vbAux = False
caracteres:
  If vbAux Then
    vOrdemBusca = 2
  End If

  ShowPopup = False

  Dim qBuscaPrestador As Object
  Dim Interface As Object
  Set qBuscaPrestador = NewQuery
  Set Interface = CreateBennerObject("Procura.Procurar")

  If PRESTADOR.LocateText <> "" Then 'Se foi digitado algo

    If vOrdemBusca = 2 Then 'o que foi digitado é o nome do prestador
      'Este select foi montado de acordo com o que já era executado na interface de procura
      qBuscaPrestador.Clear
      qBuscaPrestador.Add("SELECT COUNT(1) QTDE")
      qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
      qBuscaPrestador.Add(" WHERE RECEBEDOR = 'S' ")
      qBuscaPrestador.Add("   AND Z_NOME LIKE '" + PRESTADOR.LocateText + "%'")
      qBuscaPrestador.Add("   AND HANDLE NOT IN (SELECT PRESTADOR FROM FILIAIS_PRESTADOR WHERE FILIAL = :FILIAL)")
      qBuscaPrestador.ParamByName("FILIAL").AsInteger = CurrentQuery.FieldByName("FILIAL").AsInteger
      qBuscaPrestador.Active = True

      If (qBuscaPrestador.FieldByName("QTDE").AsInteger > 1) Or (qBuscaPrestador.FieldByName("QTDE").AsInteger = 0) Then 'A busca retornou mais de um registro, ou não retornou nenhum

        vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
        vCriterio = "RECEBEDOR = 'S' AND HANDLE NOT IN (SELECT PRESTADOR FROM FILIAIS_PRESTADOR WHERE FILIAL = "+CurrentQuery.FieldByName("FILIAL").AsString+")"
        vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
        vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Prestador", False, PRESTADOR.LocateText)

      Else 'Se encontrou um prestador com o dado digitado

        qBuscaPrestador.Clear
        qBuscaPrestador.Add("SELECT HANDLE")
        qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
        qBuscaPrestador.Add(" WHERE RECEBEDOR = 'S' ")
        qBuscaPrestador.Add("   AND Z_NOME LIKE '" + PRESTADOR.LocateText + "%'")
        qBuscaPrestador.Add("   AND HANDLE NOT IN (SELECT PRESTADOR FROM FILIAIS_PRESTADOR WHERE FILIAL = :FILIAL)")
        qBuscaPrestador.ParamByName("FILIAL").AsInteger = CurrentQuery.FieldByName("FILIAL").AsInteger
        qBuscaPrestador.Active = True

        vHandle = qBuscaPrestador.FieldByName("HANDLE").AsInteger

      End If
    Else 'Foi digitado o código do prestador
      qBuscaPrestador.Clear
      qBuscaPrestador.Add("SELECT COUNT(1) QTDE")
      qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
      qBuscaPrestador.Add(" WHERE RECEBEDOR = 'S' ")
      qBuscaPrestador.Add("   AND PRESTADOR LIKE '" + PRESTADOR.LocateText + "%'")
      qBuscaPrestador.Add("   AND HANDLE NOT IN (SELECT PRESTADOR FROM FILIAIS_PRESTADOR WHERE FILIAL = :FILIAL)")
      qBuscaPrestador.ParamByName("FILIAL").AsInteger = CurrentQuery.FieldByName("FILIAL").AsInteger
      qBuscaPrestador.Active = True

      If (qBuscaPrestador.FieldByName("QTDE").AsInteger > 1) Or (qBuscaPrestador.FieldByName("QTDE").AsInteger = 0) Then

        vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
        vCriterio = "RECEBEDOR = 'S' AND HANDLE NOT IN (SELECT PRESTADOR FROM FILIAIS_PRESTADOR WHERE FILIAL = "+CurrentQuery.FieldByName("FILIAL").AsString+")"
        vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
        vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, vCampos, vCriterio, "Prestador", False, PRESTADOR.LocateText)

      Else 'Se encontrou um prestador com o dado digitado

        qBuscaPrestador.Clear
        qBuscaPrestador.Add("SELECT HANDLE")
        qBuscaPrestador.Add("  FROM SAM_PRESTADOR")
        qBuscaPrestador.Add(" WHERE RECEBEDOR = 'S' ")
        qBuscaPrestador.Add("   AND PRESTADOR LIKE '" + PRESTADOR.LocateText + "%'")
        qBuscaPrestador.Add("   AND HANDLE NOT IN (SELECT PRESTADOR FROM FILIAIS_PRESTADOR WHERE FILIAL = :FILIAL)")
        qBuscaPrestador.ParamByName("FILIAL").AsInteger = CurrentQuery.FieldByName("FILIAL").AsInteger
        qBuscaPrestador.Active = True

        vHandle = qBuscaPrestador.FieldByName("HANDLE").AsInteger
      End If
    End If
  Else 'Se não foi digitado nada, apenas abre a interface de procura

    vColunas = "PRESTADOR|NOME|SOLICITANTE|EXECUTOR|RECEBEDOR|NAOFATURARGUIAS"
    vCriterio = "RECEBEDOR = 'S' AND HANDLE NOT IN (SELECT PRESTADOR FROM FILIAIS_PRESTADOR WHERE FILIAL = "+CurrentQuery.FieldByName("FILIAL").AsString+")"
    vCampos = "Prestador|Nome|Solicitante|Executor|Recebedor|Não faturar guias"
    vHandle = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 2, vCampos, vCriterio, "Prestador", False, PRESTADOR.LocateText)

  End If

  Set Interface = Nothing
  Set qBuscaPrestador = Nothing

  If vHandle <>0 Then
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
End Sub
