Imports System
Imports Newtonsoft.Json
Imports RestSharp

Module Program
    Sub Main(args As String())
        Console.WriteLine("Hello World!")
        Dim api As RequestApi = New RequestApi()
        Dim responseToken = api.RequestLogin()

        Dim token = New With {.token = ""}
        Dim sesion = JsonConvert.DeserializeAnonymousType(responseToken.Content, token)
        api.RequestRefOpenPay(sesion.token)

        api.EndSession(sesion.token)
    End Sub
End Module


Class RequestApi

    Private url As Uri = New Uri("https://localhost:44394/")
    Private token As String

    Public Function RequestLogin() As RestResponse
        Dim client As RestClient = New RestClient(url)
        Dim json As String = "{""Name"": ""FACTUTACIONIT"", ""Password"": ""XRpDGANzd1""}"
        Dim request = New RestRequest(Method.POST)
        request.Resource = "acceso/login"
        request.AddHeader("Content-Type", "application/json")
        request.AddParameter("application/json", json, ParameterType.RequestBody)
        Dim respuesta As RestResponse = CType(client.Execute(request), RestResponse)
        Return respuesta
    End Function

    Public Function EndSession(token As String) As RestResponse

        Dim client As RestClient = New RestClient(url)
        Dim request = New RestRequest(Method.GET)
        request.Resource = "acceso/logout"
        request.AddHeader("Content-Type", "application/json")
        request.AddHeader("token", token)
        Dim respuesta As RestResponse = CType(client.Execute(request), RestResponse)
        Return respuesta
    End Function

    Public Function RequestRefOpenPay(token As String) As RestResponse

        Dim client As RestClient = New RestClient(url)
        'Dim json As String = JsonConvert.SerializeObject(Data)
        Dim request = New RestRequest(Method.GET)
        request.Resource = "facturacion/ListadoFacturas"
        request.AddHeader("Content-Type", "application/json")
        request.AddHeader("token", token)
        'request.AddParameter("application/json", json, ParameterType.RequestBody)
        Dim respuesta As RestResponse = CType(client.Execute(request), RestResponse)
        Return respuesta
    End Function
End Class
