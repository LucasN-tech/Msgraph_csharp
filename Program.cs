using Microsoft.Graph;
using Microsoft.Graph.Applications.Item.Owners.GraphUser;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using System.Web.Http;
using Newtonsoft.Json.Linq;

class Program
{
    static async Task Main()
    {
        var clientID = "YOUR_CLIENT_ID";
        var tenant = "YOUR_TENANT_ID";
        var clientSecret = "YOUR_CLIENT_SECRET";
        var accessToken = await GetAccessToken(clientID, tenant, clientSecret);

        if (!string.IsNullOrEmpty(accessToken))
        {
            var isValidToken = await ValidateToken(accessToken);

            if (isValidToken)
            {
                Console.WriteLine("Token is valid.");
            }
            else
            {
                Console.WriteLine("Token is invalid.");
            }
        }

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    static async Task<string> GetAccessToken(string clientID, string tenant, string clientSecret)
    {
        using (var httpClient = new HttpClient())
        {
            var requestBody = $"client_id={clientID}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret={clientSecret}&grant_type=client_credentials";
            Console.WriteLine(requestBody);
            var content = new StringContent(requestBody, Encoding.UTF8, "application/x-www-form-urlencoded");
            var response = await httpClient.PostAsync($"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", content);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                var token = result.Substring(result.IndexOf("access_token\":\"") + 15);
                token = token.Substring(0, token.IndexOf("\""));
                return token;
            }

            Console.WriteLine("Failed to retrieve access token.");
            return null;
        }
    }

    static async Task<bool> ValidateToken(string token)
    {
        using (var httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await httpClient.GetAsync("https://login.microsoftonline.com/common/oauth2/nativeclient"); // Substitua com a URL de validação desejada

            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine("Atividade ValidateToken: Token is valid.");
                return true;
            }

            Console.WriteLine("Atividade ValidateToken: Token is invalid.");
            return false;
        }
    }

    static async Task<HttpResponseMessage> GetUsers(string token)
    {
        using (var httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            try
            {
                var url = "https://graph.microsoft.com/v1.0/users";
                var response = await httpClient.GetAsync(url);

                return response;
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine("Error in HTTP request: " + ex.Message);
                return null;
            }
        }
    }

    static async Task<HttpResponseMessage> GetEmailMessages(string token)
    {
        using (var httpClient = new HttpClient())
        {
            var userEmail = "email from";
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = "https://graph.microsoft.com/v1.0/users/{userID}/messages";

            try
            {
                var response = await httpClient.GetAsync(url);
                return response;
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine("Error in HTTP request: " + ex.Message);
                return null;
            }
        }
    }

    static async Task<HttpResponseMessage> SendEmail(string token)
    {
        using (var httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = "https://graph.microsoft.com/v1.0/users/{Email_from}/sendMail";

            var email = new
            {
                message = new
                {
                    subject = "SUBTITULO",
                    body = new
                    {
                        contentType = "Text",
                        content = "Olá, este é um exemplo de email!"
                    },
                    toRecipients = new[]
                    {
                    new
                    {
                        emailAddress = new
                        {
                            address = "cc"
                        }
                    }
                }
                }
            };

            var jsonEmail = JsonConvert.SerializeObject(email);
            var content = new StringContent(jsonEmail, Encoding.UTF8, "application/json");

            try
            {
                var response = await httpClient.PostAsync(url, content);
                Console.WriteLine("enviado");
                return response;
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine("Error in HTTP request: " + ex.Message);
                return null;
            }
        }
    }
}