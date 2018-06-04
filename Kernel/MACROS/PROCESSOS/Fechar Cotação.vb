'HASH: 9FB34D692758D3DD1D3AB0E057E586E4
Sub Main
    Dim obj As Object
    Set obj = CreateBennerObject("bsportalweb.web")
    obj.FecharLicitacao(CurrentSystem,SessionVar("LICITACAO"))
    Set obj = Nothing
  End sub
