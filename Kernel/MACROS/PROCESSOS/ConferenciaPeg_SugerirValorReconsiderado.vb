'HASH: B33F804597F49879A46D3833BEA6E547
Option Explicit

Sub Main()


On Error GoTo erro:
	Dim vHGuiaEvento As Long
	Dim vHGlosa As Long
	Dim vfQtdReconsiderada As Double
	Dim vsMsgRetorno As String
	Dim vfValorRecSugerido As Double

	vHGuiaEvento = CLng( ServiceVar("pHGuiaEvento") )
	vHGlosa = CLng( ServiceVar("pHGlosa") )
	vfQtdReconsiderada = CDbl(ServiceVar("pfQtdReconsiderada") )

    Dim interface As Object
    Set interface=CreateBennerObject("SAMPEG.processar")
    vfValorRecSugerido = interface.SugerirValorReconsiderado(CurrentSystem, vHGuiaEvento, vHGlosa, vfQtdReconsiderada)
    Set interface=Nothing

    ServiceVar("psMsgRetorno") = CStr(vsMsgRetorno)
	ServiceVar("pfValorRecSugerido") = CStr(vfValorRecSugerido)


Exit Sub

erro:
    ServiceVar("psMsgRetorno") = CStr(Err.Description)


End Sub
