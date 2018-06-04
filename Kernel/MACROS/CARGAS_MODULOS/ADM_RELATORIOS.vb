'HASH: 4158863A3FBBD54A2936B503E8F78642
Public Sub EXCLUIRTODOS_OnClick()
  Dim sql As Object
  Set sql =NewQuery

  If Not MsgBox(("Esta operação excluirá todos os relatórios da base. Deseja continuar?"),vbYesNo,"Atenção")=vbYes Then
    Exit Sub
  End If
  If Not InTransaction Then
    StartTransaction
  End If
  sql.Clear
  sql.Add("UPDATE R_RELATORIOS Set DETALHE=Null")
  sql.ExecSQL
  sql.Clear
  sql.Add("UPDATE R_RELATORIOS Set CABECALHO=Null")
  sql.ExecSQL
  sql.Clear
  sql.Add("UPDATE R_RELATORIOS Set RODAPE=Null")
  sql.ExecSQL
  sql.Clear
  sql.Add("UPDATE R_RELATORIOS Set SUMARIO=Null")
  sql.ExecSQL
  sql.Clear
  sql.Add("UPDATE R_RELATORIOS Set TITULO=Null")
  sql.ExecSQL
  sql.Clear
  sql.Add("UPDATE SAM_PARAMETROSBENEFICIARIO")
  sql.Add("Set RELATORIOCARTAO = Null")
  sql.ExecSQL
  sql.Clear
  sql.Add("UPDATE R_DETALHES Set RELATORIO=Null")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_RELATORIOUSUARIOS")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_DETALHECAMPOS")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_QUEBRASDETALHE")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_DETALHEDETALHES")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_DETALHES")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_GRUPORELATORIOS")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_RELATORIOS_CORRESP")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_RELATORIOS_TIPOCORRESP")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_RELATORIOIMPRESSOES")
  sql.ExecSQL
  sql.Clear
  sql.Add("DELETE FROM R_RELATORIOS")
  sql.ExecSQL
  If InTransaction Then
    Commit
  End If

End Sub

Public Sub IMPORTAR_OnClick()
Dim ImpObj As Object
 Set ImpObj =CreateBennerObject("CS.RelImportar")
 ImpObj.Exec()
 Set ImpObj =Nothing
End Sub
Public Sub EXPORTAR_OnClick()
Dim ImpObj As Object
 Set ImpObj =CreateBennerObject("CS.RelExportar")
ImpObj.Exec()
 Set ImpObj =Nothing
End Sub
Public Sub RELATORIOAGENDAMENTO_OnClick()
Dim ImpObj As Object
 Set ImpObj =CreateBennerObject("CS.RelatoriosAgendados")
 ImpObj.Exec(CurrentSystem)
 Set ImpObj =Nothing
End Sub
