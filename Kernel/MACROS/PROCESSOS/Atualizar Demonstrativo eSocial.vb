'HASH: F676975F44B7C86660BDEB2B0DD1B4E7

'eSocial - Atualizar Demonstrativo

Public Sub Main
 Dim business As CSBusinessComponent
 Set business = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.Rotinas.Demonstrativo.Processos, Benner.Saude.Financeiro.Business")
 business.Execute("AtualizarDemonstrativoESocial")
 Set Processos = Nothing
End Sub
