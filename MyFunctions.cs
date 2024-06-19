using System.Text;
using ExcelDna.Documentation;
using ExcelDna.Integration;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

public static class MyFunctions
{
    private static string DEFAULT_MODEL = "qwen2:1.5b-instruct-q4_K_M";
    private static string SYSTEM_PROMPT = "You are a helpful assistant who loves numbers.";

    [ExcelFunctionDoc(Description = "Generate a response from a LLM", HelpTopic = "Ask-LLM-AddIn.chm!1001")]
    public static object ASK(
        [ExcelArgument(Name = "prompt", Description = "a question/context for the LLM")]
        string UserInput)
    {
        return ExcelAsyncUtil.Run(nameof(PostNetRequest), new object[] { UserInput }, () => PostNetRequest(UserInput));
    }

    [ExcelFunctionDoc(Description = "Interprets data with the help of a prompt", HelpTopic = "Ask-LLM-AddIn.chm!1002")]
    public static object INTERPRET(
        [ExcelArgument(Name = "prompt", Description = "context for the LLM")]
        string Prompt,
        [ExcelArgument(Name = "user_input", Description = "excel data that is appended to the prompt")]
        object[,] UserInputs)
    {
        string ArrayRep = Prompt + ": ";
        try
        {
            for (int i = 0; i < UserInputs.GetLength(0); i++)
            {
                ArrayRep += "[";
                for (int j = 0; j < UserInputs.GetLength(1); j++)
                {
                    string input = UserInputs[i, j].ToString() ?? throw new NullReferenceException("Excel input is null!");
                    if (input == "ExcelDna.Integration.ExcelEmpty")
                    {
                        input = "0";
                    }
                    ArrayRep += input;
                    ArrayRep += (j == UserInputs.GetLength(1) - 1) ? "" : ",";
                }
                ArrayRep += "], ";
            }
            if (ArrayRep.Contains(','))
            {
                ArrayRep = ArrayRep.Substring(0, ArrayRep.LastIndexOf(','));
            }
        }
        catch (Exception e)
        {
            return "ERROR: " + e.Message;
        }
        return ExcelAsyncUtil.Run(nameof(PostNetRequest), new object[] { ArrayRep }, () => PostNetRequest(ArrayRep));
    }

    private static string PostNetRequest(string UserInput)
    {
        try
        {
            // Environment variable for the model
            string? env = Environment.GetEnvironmentVariable("EXCEL_MODEL", EnvironmentVariableTarget.User);
            if (env == null)
            {
                Environment.SetEnvironmentVariable("EXCEL_MODEL", DEFAULT_MODEL, EnvironmentVariableTarget.User);
                env = DEFAULT_MODEL;
            }

            // Requires Ollama
            var requestUrl = "http://localhost:11434/api/chat";
            var requestPayload = new
            {
                model = env,
                temperature = 0.0,
                stream = false,
                messages = new[]
                {
                    new { role = "system", content = SYSTEM_PROMPT },
                    new { role = "user", content = UserInput }
                }
            };

            using (var client = new HttpClient())
            {
                var requestContent = new StringContent(
                JsonConvert.SerializeObject(requestPayload),
                Encoding.UTF8,
                "application/json");

                var response = client.PostAsync(requestUrl, requestContent).Result;
                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"HTTP request failed with status code {response.StatusCode}!");
                }

                var jsonResponse = response.Content.ReadAsStringAsync().Result;
                var jsonObject = JObject.Parse(jsonResponse);
                JToken Content = (jsonObject["message"] ?? throw new NullReferenceException("Cannot access \"message\" in JSON string"))["content"] ?? throw new NullReferenceException("Cannot access \"content\" in JSON string");

                return Content.ToString();
            }
        }
        catch (Exception e)
        {
            return "ERROR: " + e.Message;
        }
    }
}