Sub PublicApiServiceBrSoUps()

Dim sEnv As Variant
Dim sURL As String
Dim xmlhtp As Object, xmlDoc As Object, b

sURL = "http://sbr-test.so-ups.ru:8091/PublicApi/PublicApiService.svc"


    sEnv = "<?xml version=""1.0"" encoding=""utf-8""?>"
    sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:pub=""http://www.armd.ru/soft/dssi/SOEES/SBR/Web/Api/PublicApi"">"
    sEnv = sEnv & "<soapenv:Header/>"
    sEnv = sEnv & "<soapenv:Body>"
    sEnv = sEnv & "<pub:GetVsvgoZoneFlowData>"
    sEnv = sEnv & "<pub:date>2022-10-02</pub:date>"
    sEnv = sEnv & "</pub:GetVsvgoZoneFlowData>"
    sEnv = sEnv & "</soapenv:Body>"
    sEnv = sEnv & "</soapenv:Envelope>"

    Set xmlhtp = CreateObject("Microsoft.XMLHTTP")
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")

    b = Len(sEnv)

    With xmlhtp
        .Open "POST", sURL, False
        .setRequestHeader "accept-Encoding", "gzip,deflate"
        .setRequestHeader "Content-Length", b
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "soapAction", "http://www.armd.ru/soft/dssi/SOEES/SBR/Web/Api/PublicApi/IPublicApiService/GetVsvgoZoneFlowData"
        .setRequestHeader "Host", "sbr-test.so-ups.ru:8091"
        .setRequestHeader "Connection", "Keep-Alive"
        .setRequestHeader "User-Agent", "Apache-HttpClient/4.5.5 (Java/16.0.1)"
        .send sEnv

        xmlDoc.LoadXML .responseText
        MsgBox .responseText
    End With
    
End Sub
