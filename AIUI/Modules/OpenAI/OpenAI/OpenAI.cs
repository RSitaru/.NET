namespace OpenAI;

using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using OpenAI;

class ChatGPT
{
    private const string OpenAIApiKey = "YOUR_OPENAI_API_KEY";

    public static async Task<string> ProcessDocxAsync(string prompt, byte[] docxFile)
    {
        try
        {
            // Create a new instance of the OpenAI API client
            OpenAI.ApiKey = OpenAIApiKey;

            // Convert byte array to text
            string text = Encoding.UTF8.GetString(docxFile);

            // Split text into chunks
            const int ChunkSize = 4000;
            var textChunks = Enumerable.Range(0, text.Length / ChunkSize)
                .Select(i => text.Substring(i * ChunkSize, ChunkSize))
                .ToList();

            var results = new List<string>();

            foreach (var chunk in textChunks)
            {
                var messages = new List<OpenAI.ChatCompletionMessage>
                {
                    new OpenAI.ChatCompletionMessage
                    {
                        Role = "user",
                        Content = $"{prompt} {chunk}"
                    }
                };

                var response = await OpenAI.ChatCompletion.CreateAsync(
                    model: "gpt-3.5-turbo",
                    messages: messages
                );

                var result = response.Choices[0]?.Message?.Content?.Trim();
                if (!string.IsNullOrEmpty(result))
                {
                    results.Add(result);
                }
            }

            // Combine the results into a single string
            var processedText = string.Join("\n", results);

            return processedText;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing the DOCX file: {ex.Message}");
            return null;
        }
    }
}
