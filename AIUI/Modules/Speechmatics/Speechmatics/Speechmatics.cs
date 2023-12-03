namespace Speechmatics;
using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

class Speechmatics
{
    private const string SpeechmaticsApiKey = "YOUR_SPEECHMATICS_API_KEY";

    public static async Task<string> UploadAudioForTranscriptionAsync(byte[] audioFile, string filename)
    {
        try
        {
            // Create an HttpClient instance
            using var client = new HttpClient();

            // Set the API key in the request headers
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", SpeechmaticsApiKey);

            // Create a form content for uploading the audio file
            var formContent = new MultipartFormDataContent();
            formContent.Add(new ByteArrayContent(audioFile), "data_file", filename);

            // Send a POST request to initiate transcription
            var response = await client.PostAsync("https://asr.api.speechmatics.com/v2/jobs/", formContent);

            if (response.IsSuccessStatusCode)
            {
                return "Audio file uploaded and being processed. Check back later for the transcription.";
            }
            else
            {
                return $"Error uploading audio file: {response.StatusCode} - {response.ReasonPhrase}";
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error uploading audio file: {ex.Message}");
            return null;
        }
    }

    public static async Task<string> GetTranscriptionAsync(string jobId)
    {
        try
        {
            // Create an HttpClient instance
            using var client = new HttpClient();

            // Set the API key in the request headers
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", SpeechmaticsApiKey);

            // Send a GET request to retrieve the transcription
            var response = await client.GetAsync($"https://asr.api.speechmatics.com/v2/jobs/{jobId}/transcript?format=txt");

            if (response.IsSuccessStatusCode)
            {
                var transcription = await response.Content.ReadAsStringAsync();
                return transcription;
            }
            else
            {
                return $"Error retrieving transcription: {response.StatusCode} - {response.ReasonPhrase}";
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving transcription: {ex.Message}");
            return null;
        }
    }
}
