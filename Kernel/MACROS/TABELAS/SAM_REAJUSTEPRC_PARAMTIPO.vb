'HASH: 615A3F7AFBE2F9BD6ECEE31E141F3C8E
'Macro: SAM_REAJUSTEPRC_PARAMTIPO
Attribute VB_Name = "Module1"
Dim vEstadodaTabela As Long

Option Explicit

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vColunas As String
  Dim vCriterios As String
  Dim vHandle As Long
  Dim vCampos As String
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TABUS.DESCRICAO"

  vCriterios = ""
  vCampos = "Descrição"

  vHandle = interface.exec(CurrentSystem, "SAM_TABUS", vColunas, 1, vCampos, vCriterios, "Tabela de US ", False, "")
  CurrentQuery.FieldByName("TABELAUS").AsInteger = vHandle
  Set Interface = Nothing
  ShowPopup = False
End Sub

Public Sub TABLE_AfterPost()
  ExcluirFilhos
  If vEstadodaTabela = 2 Then
    MsgBox "Atenção! Você deve GERAR novamente o reajuste!"
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  ExcluirFilhos
  RefreshNodesWithTable("SAM_REAJUSTEPRC_PARAM")
End Sub

Public Sub ExcluirFilhos
  Dim PAI As Object
  Dim DEL As Object
  Set PAI = NewQuery
  Set DEL = NewQuery
  PAI.Clear
  PAI.Add("SELECT HANDLE, PRESTADOR, ASSOCIACAO, ESTADO, MUNICIPIO, REDERESTRITA FROM SAM_REAJUSTEPRC_PARAM WHERE HANDLE =:HANDLE")
  PAI.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("REAJUSTEPRCPARAM").AsInteger
  PAI.Active = True

  If Not PAI.FieldByName("PRESTADOR").IsNull Then
    Excluir "PRESTADOR", PAI.FieldByName("HANDLE").AsInteger
  End If
  If Not PAI.FieldByName("ASSOCIACAO").IsNull Then
    Excluir "PRESTADOR", PAI.FieldByName("HANDLE").AsInteger
  End If
  If Not PAI.FieldByName("ESTADO").IsNull And Pai.FieldByName("MUNICIPIO").IsNull Then
    Excluir "ESTADO", PAI.FieldByName("HANDLE").AsInteger
  End If
  If Not PAI.FieldByName("ESTADO").IsNull And Not Pai.FieldByName("MUNICIPIO").IsNull Then
    Excluir "MUNICIPIO", PAI.FieldByName("HANDLE").AsInteger
  End If
  If Not PAI.FieldByName("REDERESTRITA").IsNull Then
    Excluir "REDE", PAI.FieldByName("HANDLE").AsInteger
  End If
  PAI.Active = False
  Set PAI = Nothing
End Sub

Public Sub Excluir(T As String, Handle As Long)
  Dim T2 As String
  Dim DEL As Object
  Set DEL = NewQuery

  Select Case CurrentQuery.FieldByName("TIPODOREAJUSTE").AsString
    Case "D"
      T2 = "DOT"
    Case "R"
      T2 = "REG"
    Case "A"
      T2 = "AN"
    Case "S"
      T2 = "SL"
  End Select

  DEL.Add("DELETE FROM SAM_REAJUSTEPRC_" + T + "_" + T2 + " WHERE REAJUSTEPRCPARAM = :REAJUSTEPRCPARAM And PARAMTIPO = :PARAMTIPO")

  DEL.ParamByName("REAJUSTEPRCPARAM").Value = Handle
  DEL.ParamByName("PARAMTIPO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DEL.ExecSQL
  Set DEL = Nothing
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  vEstadodaTabela = CurrentQuery.State
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > _
                              CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    MsgBox("Data INICIAL não pode ser maior que a data FINAL", "Reajuste de Preço")
    CanContinue = False
  ElseIf CurrentQuery.FieldByName("NOVAVIGENCIA").AsDateTime <= _
                                    CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    MsgBox("NOVA VIGÊNCIA deve ser maior que a data FINAL", "Reajuste de Preço")
    CanContinue = False
  End If
End Sub

