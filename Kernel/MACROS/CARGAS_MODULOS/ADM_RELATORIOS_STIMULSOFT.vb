'HASH: 48474430286C85A01ADBA56584EAA5CC
 
Public Sub IMPORTAR_OnClick() 
Dim ImpObj As Object 
 Set ImpObj = CreateBennerObject("CS.RelImportar") 
 ImpObj.Exec 
 Set ImpObj = Nothing 
End Sub 
 
Public Sub EXPORTAR_OnClick() 
Dim ImpObj As Object 
 Set ImpObj = CreateBennerObject("CS.RelExportar") 
ImpObj.Exec 
 Set ImpObj = Nothing 
End Sub 
