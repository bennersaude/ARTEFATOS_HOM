'HASH: 4BCBBA7434D9462ACFCAED1FD1BE9B3B


Public Sub TABLE_AfterScroll()

  CARATERSOLICITACAO.LocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO")) 'SMS 103749 - 04/11/2008 - Evandro Zeferino

End Sub


Public Sub TABLE_NewRecord()

  CARATERSOLICITACAO.LocalWhere = " VERSAOTISS = " + CStr(RecordHandleOfTable("TIS_VERSAO")) 'SMS 103749 - 04/11/2008 - Evandro Zeferino

End Sub

