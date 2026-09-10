using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace renamePDF_v2
{
    public class DeepSeekClient
    {
        private const string API_URL = "https://api.deepseek.com/chat/completions";
        private const string MODEL = "deepseek-v4-flash";
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;

        public DeepSeekClient(string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new ArgumentException("DeepSeek API Key không được để trống.");
            _apiKey = apiKey;
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromMinutes(5);
        }

        // =========================================================
        // Gọi DeepSeek
        // =========================================================

        public async Task<string?> ChatAsync(
            string systemPrompt,
            string userPrompt)
        {
            if (string.IsNullOrWhiteSpace(systemPrompt))
                throw new ArgumentException(
                    "System prompt không được để trống.");

            if (string.IsNullOrWhiteSpace(userPrompt))
                throw new ArgumentException(
                    "User prompt không được để trống.");

            var requestBody = new
            {
                model = MODEL,

                messages = new object[]
                {
                    new
                    {
                        role = "system",
                        content = systemPrompt
                    },

                    new
                    {
                        role = "user",
                        content = userPrompt
                    }
                },

                response_format = new
                {
                    type = "json_object"
                },

                temperature = 0.0,

                max_tokens = 500,

                stream = false
            };

            string jsonRequest =
                JsonSerializer.Serialize(requestBody);

            using var request =
                new HttpRequestMessage(
                    HttpMethod.Post,
                    API_URL);

            request.Headers.Authorization =
                new AuthenticationHeaderValue(
                    "Bearer",
                    _apiKey);

            request.Content =
                new StringContent(
                    jsonRequest,
                    Encoding.UTF8,
                    "application/json");

            using HttpResponseMessage response =
                await _httpClient.SendAsync(request);

            string responseText =
                await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception(
                    $"DeepSeek API lỗi " +
                    $"{(int)response.StatusCode} " +
                    $"{response.StatusCode}\n\n" +
                    responseText);
            }

            DeepSeekResponse? result;

            try
            {
                result = JsonSerializer.Deserialize<DeepSeekResponse>(
                        responseText);
            }
            catch (Exception ex)
            {
                throw new Exception(
                    "Không đọc được response từ DeepSeek.\n\n" +
                    responseText,
                    ex);
            }

            string? content =
                result?
                    .Choices?
                    .FirstOrDefault()?
                    .Message?
                    .Content;

            if (string.IsNullOrWhiteSpace(content))
            {
                return null;
            }

            return CleanJson(content);
        }

        // =========================================================
        // Làm sạch JSON
        // =========================================================

        private string CleanJson(string input)
        {
            input = input.Trim();

            if (input.StartsWith("```"))
            {
                int firstNewLine =
                    input.IndexOf('\n');

                int lastFence =
                    input.LastIndexOf("```");

                if (firstNewLine >= 0 &&
                    lastFence > firstNewLine)
                {
                    return input.Substring(
                        firstNewLine + 1,
                        lastFence - firstNewLine - 1)
                        .Trim();
                }
            }

            return input;
        }

        // =========================================================
        // Response models
        // =========================================================

        private class DeepSeekResponse
        {
            [JsonPropertyName("choices")]
            public List<DeepSeekChoice>? Choices { get; set; }
        }

        private class DeepSeekChoice
        {
            [JsonPropertyName("message")]
            public DeepSeekMessage? Message { get; set; }

            [JsonPropertyName("finish_reason")]
            public string? FinishReason { get; set; }
        }

        private class DeepSeekMessage
        {
            [JsonPropertyName("content")]
            public string? Content { get; set; }
        }
    }
}
