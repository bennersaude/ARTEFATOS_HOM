'HASH: 9FDF161EBAEA823C7F6BD00E15163215
Option Explicit

'#USES "*CriaTabelaTemporariaSqlServer"
'#Uses "*bsShowMessage"

Public Sub Main

  Dim vAbriuTransacao As Boolean
  vAbriuTransacao = False

  On Error GoTo erro
  If Not InTransaction Then
    StartTransaction
    vAbriuTransacao = True
  End If

  Dim viContador1 As Long
  NewCounter("AUTORIZADORSP", 0, 1, viContador1)

  Dim viContador2 As Long
  NewCounter("AUTORIZADORSP", 0, 1, viContador2)

  Dim qPegsAProcessar As Object
  Set qPegsAProcessar = NewQuery
  qPegsAProcessar.Clear
  qPegsAProcessar.Add("UPDATE SAM_PEG_PROCESSOLOTE SET CHAVE = :CHAVE WHERE SITUACAO = '1'")
  qPegsAProcessar.ParamByName("CHAVE").AsInteger = viContador1
  qPegsAProcessar.ExecSQL
  Set qPegsAProcessar = Nothing

  If InStr(SQLServer, "MSSQL")>0 Then
    CriaTabelaTemporariaSqlServer
  End If

  Dim SPP As BStoredProc
  Set SPP = NewStoredProc
  SPP.AutoMode = True
  SPP.Name = "BSPROPEG_MUDARFASELOTE"
  SPP.AddParam("P_CHAVE",ptInput, ftInteger)
  SPP.AddParam("P_CHAVEAUX",ptInput, ftInteger)
  SPP.AddParam("P_USUARIO",ptInput, ftInteger)
  SPP.ParamByName("P_CHAVE").AsInteger = viContador1
  SPP.ParamByName("P_CHAVEAUX").AsInteger = viContador2
  SPP.ParamByName("P_USUARIO").AsInteger = CurrentUser
  SPP.ExecProc

  If vAbriuTransacao Then
    Commit
  End If

  Exit Sub

  Erro :

  If vAbriuTransacao Then
    Rollback
  End If

  bsShowMessage("Erro ao executar Mudança de Fase em Lote" + Err.Description, "I")

End Sub
