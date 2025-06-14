using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text;

namespace WordFixer.Console.Service
{
    public class OpenAIService
    {
        private readonly string _apiKey;
        private readonly HttpClient _httpClient;

        public OpenAIService(string apiKey)
        {
            _apiKey = apiKey;
            _httpClient = new HttpClient{ Timeout = TimeSpan.FromMinutes(5) };
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
        }

        public async Task<TextoExtraido> CorrigirTextosAsync(string texto)
        {
            try
            {
                string prompt = CriarPrompt(texto);

                var requestBody = new
                {
                    model = "gpt-4", // ou "gpt-3.5-turbo
                    messages = new[]
                    {
                        new { role = "system", content = @$"Você é um corretor de texto especializado em português. Corrija erros gramaticais, ortográficos e de concordância verbal. Não altere o estilo ou o significado do texto. Retorne no formato JSON e o novo texto deve estar sem quebra de linhas '\n': {{'TextoNovo': 'texto corrigido', 'PalavraComErro': ''}}." },
                        new { role = "user", content = texto }
                    }
                };

                string jsonBody = JsonConvert.SerializeObject(requestBody);

                var response = await _httpClient.PostAsync(
                    "https://api.openai.com/v1/chat/completions",
                    new StringContent(jsonBody, Encoding.UTF8, "application/json")
                );

                string responseBody = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"Erro na chamada OpenAI: {response.StatusCode}. Detalhes: {responseBody}");
                }

                var resposta = JsonConvert.DeserializeObject<dynamic>(responseBody);
                if (resposta == null)
                {
                    throw new InvalidOperationException("A resposta da API está vazia ou inválida.");
                }

                string? mensagem = resposta?.choices?[0]?.message?.content;

                if (string.IsNullOrEmpty(mensagem))
                {
                    throw new InvalidOperationException("A resposta da API está vazia ou inválida.");
                }

                var textoExtraido = JsonConvert.DeserializeObject<TextoExtraido>(mensagem);
                if (textoExtraido == null)
                {
                    throw new InvalidOperationException("Falha ao desserializar a resposta da API para o tipo TextoExtraido.");
                }

                return textoExtraido;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string CriarPrompt(string texto)
        {
            return @$"Corrija os textos abaixo e devolva no formato: 
                    {texto}
                    ";
        }
    }

    public class TextoExtraido
    {
        public string TextoNovo { get; set; }
        public string PalavraComErro { get; set; }
    }
}
