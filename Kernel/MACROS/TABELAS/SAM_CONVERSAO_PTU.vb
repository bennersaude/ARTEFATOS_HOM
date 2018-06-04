'HASH: 2A582572E185598B4F52795A7DE2FBA8


Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim sql As Object


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, "")

  If vHandle <>0 Then
    Set sql = NewQuery
    sql.Clear
    sql.Active = False
    sql.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")
    sql.ParamByName("HEVENTO").Value = vHandle
    sql.Active = True
    CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = sql.FieldByName("ESTRUTURA").Value


    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim sql As Object


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, "")

  If vHandle <>0 Then
    Set sql = NewQuery
    sql.Clear
    sql.Active = False
    sql.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")
    sql.ParamByName("HEVENTO").Value = vHandle
    sql.Active = True
    CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = sql.FieldByName("ESTRUTURA").Value


    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
  Set interface = Nothing

End Sub

