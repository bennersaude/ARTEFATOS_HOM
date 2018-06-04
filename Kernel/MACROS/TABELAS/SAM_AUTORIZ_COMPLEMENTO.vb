'HASH: 0CE32200F5BD7EA5F31203231D91496D
Public Sub TABLE_AfterScroll()
  WriteBDebugMessage("SAM_AUTORIZ_COMPLEMENTO.TABLE_AfterScroll - Início")
  'No WES 2006 o LocalWhere proveniente de entidade especializada não sensibiliza = WebLocalWhere automaticamente
  EVENTO.WebLocalWhere = EVENTO.LocalWhere
  DATASOLICITACAO.ReadOnly = True
  WriteBDebugMessage("SAM_AUTORIZ_COMPLEMENTO.TABLE_AfterScroll - Fim")
End Sub
