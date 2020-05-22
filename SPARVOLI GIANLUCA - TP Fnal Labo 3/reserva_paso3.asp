<body>
    <%
     Response.Write("Apellido y nombre: " & Request.form("apellido_nombre"))
     Response.Write("<br />Destino: " & Request.form("destino"))
     Response.Write("<br />Hora Partida: " & Request.form("hora"))
     Response.Write("<br />Cantidad Pasajeros: " & Request.form("cantsapajeros"))
     Response.Write("<br />Reserva Traslado: " )
        If Len(Request.form("ProvedorTaxi")) > 1 Then
            Response.Write("SI")
            Response.Write("<br />" & Request.form("ProvedorTaxi"))
        Else
            Response.Write("NO")
        End If
     
     Response.Write("<br />" & Request.form("ASIENTOS_SELECCIONADOS")) 
    %>
</body>
