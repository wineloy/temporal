Imports DarkSoft.SQL

Module Module1

    Sub Main()
        Console.WriteLine("Poliza IT vesión Pro")

        Dim poliza As Balance = New Balance()
        Dim listado As List(Of ResultadosPoliza)

        listado = poliza.Poliza("2021/08/04", "2021/08/04", "ILOXTELECOM")

        For Each resultado In listado
            Console.WriteLine(resultado.Descripcion)
            Console.WriteLine(resultado.Cantidad)

        Next

        Console.WriteLine("--------------------------------------------------------------")
        listado = poliza.Poliza(Today.AddDays(-5), Today.AddDays(-5), "ILOXTELECOM")

        For Each resultado In listado
            Console.WriteLine(resultado.Descripcion)
            Console.WriteLine(resultado.Cantidad)

        Next

        Console.ReadKey()

    End Sub

End Module


Public Class Balance

    Private _connection As Conexion = New Conexion("WEBTV", "192.168.60.150", "DeSoftv", "1975huli")


    Public Function Poliza(Fecha_inicial As DateTime, Fecha_final As DateTime, sistema As String) As List(Of ResultadosPoliza)

        'Lista de facturas
        Dim ResultListado As List(Of ResultadosPoliza) = New List(Of ResultadosPoliza)

        'Query dinamica base de datos
        Dim sqlPoliza As String =
            "SELECT 'CAJAS', SUM(monto) AS 'CAJAS' FROM Mayoristas.DBO.Historial_Pagos where FECHA BETWEEN " & "'" & Fecha_inicial.ToString("yyyy-dd-MM") & "'" & " AND " & "'" & Fecha_final.ToString("yyyy-dd-MM") & " 23:59:59'" & " and upper(estatus) like '%200OK%' AND tipo NOT IN('BONIFICACIÓN','CARGO','CARGO SALDO') AND UPPER(concepto)<>'CAMBIO PLAN SALDO'
            UNION 
            -- FACTURAS OPENPAY
            SELECT descripcion, SUM(MO.importe) AS 'TOTAL' FROM movimientosFacturaOpenpay MO inner join FacturasOpenPay FO ON (MO.id_factura = FO.id_factura)  WHERE FO.sistema_factura = " & "'" & sistema & "'" & " AND FO.estatus_factura = 2 AND FO.fecha_pago BETWEEN " & "'" & Fecha_inicial.ToString("yyyy-dd-MM") & "'" & " AND " & "'" & Fecha_final.ToString("yyyy-dd-MM") & " 23:59:59'" & " group by descripcion
            UNION 
            select 'iva', SUM(MO.impuestoIVA) AS 'Total IVA' FROM movimientosFacturaOpenpay MO inner join FacturasOpenPay FO ON (MO.id_factura = FO.id_factura)  WHERE FO.sistema_factura = " & "'" & sistema & "'" & " AND FO.estatus_factura = 2 AND FO.fecha_pago BETWEEN " & "'" & Fecha_inicial.ToString("yyyy-dd-MM") & "'" & " AND " & "'" & Fecha_final.ToString("yyyy-dd-MM") & " 23:59:59 '" & "
            UNION
            SELECT 'IEPS', SUM (MO.impuestoIEPS) as 'Total IEPS' FROM  movimientosFacturaOpenpay MO inner join FacturasOpenPay FO ON (MO.id_factura = FO.id_factura)  WHERE FO.sistema_factura = " & "'" & sistema & "'" & " AND FO.estatus_factura = 2 AND FO.fecha_pago BETWEEN " & "'" & Fecha_inicial.ToString("yyyy-dd-MM") & "'" & " AND " & "'" & Fecha_final.ToString("yyyy-dd-MM") & " 23:59:59" & "'"

        Dim dtPoliza As DataTable = _connection.ConsultarDT(sqlPoliza)
        If dtPoliza IsNot Nothing AndAlso dtPoliza.Rows.Count > 0 Then

            Dim resultado As ResultadosPoliza
            For index = 0 To dtPoliza.Rows.Count - 1
                resultado = New ResultadosPoliza With {
                    .Descripcion = dtPoliza.Rows(index)(0).ToString(),
                    .Cantidad = Decimal.Parse(dtPoliza.Rows(index)(1).ToString())
                }

                ResultListado.Add(resultado)
            Next

        End If

        Return ResultListado

    End Function
End Class

Public Class ResultadosPoliza
    Property Descripcion As String
    Property Cantidad As Decimal
End Class
            
            
SELECT 'CAJAS', SUM(monto) AS 'CAJAS' FROM Mayoristas.DBO.Historial_Pagos where FECHA BETWEEN '2021-04-08' AND '2021-04-08 23:59:59' and upper(estatus) like '%200OK%' AND tipo NOT IN('BONIFICACIÓN','CARGO','CARGO SALDO') AND UPPER(concepto)<>'CAMBIO PLAN SALDO'
UNION 
-- FACTURAS OPENPAY
SELECT descripcion, SUM(MO.importe) AS 'TOTAL' FROM movimientosFacturaOpenpay MO inner join FacturasOpenPay FO ON (MO.id_factura = FO.id_factura)  WHERE FO.sistema_factura = 'ILOXTELECOM' AND FO.estatus_factura = 2 AND FO.fecha_pago BETWEEN '2021-04-08' AND '2021-04-08 23:59:59' group by descripcion
UNION 
select 'iva', SUM(MO.impuestoIVA) AS 'Total IVA' FROM movimientosFacturaOpenpay MO inner join FacturasOpenPay FO ON (MO.id_factura = FO.id_factura)  WHERE FO.sistema_factura = 'ILOXTELECOM' AND FO.estatus_factura = 2 AND FO.fecha_pago BETWEEN '2021-04-08' AND '2021-04-08 23:59:59 '
UNION
SELECT 'IEPS', SUM (MO.impuestoIEPS) as 'Total IEPS' FROM  movimientosFacturaOpenpay MO inner join FacturasOpenPay FO ON (MO.id_factura = FO.id_factura)  WHERE FO.sistema_factura = 'ILOXTELECOM' AND FO.estatus_factura = 2 AND FO.fecha_pago BETWEEN '2021-04-08' AND '2021-04-08 23:59:59'
          
            
            
