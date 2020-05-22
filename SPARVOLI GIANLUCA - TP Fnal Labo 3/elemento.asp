<%
    strCp = request.QueryString("cp")

    Dim xmlLocalidades       
    Dim locNombre
    Dim locCP
    Dim proveedor
    Dim res
    res = ""
    
    Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")    
    objXMLDoc.async = False    
    objXMLDoc.load Server.MapPath("/localidades_traslados.xml")

    For Each xmlLocalidades In objXMLDoc.documentElement.selectNodes("loc")
        locNombre = xmlLocalidades.selectSingleNode("nombre").text
        locCP = xmlLocalidades.selectSingleNode("cp").text
        proveedor = xmlLocalidades.selectSingleNode("proveedor").text
        If locCp = strCp Then
                res = locNombre & "|" & proveedor
            Exit For
        End If
    Next 

    response.Write res
%>