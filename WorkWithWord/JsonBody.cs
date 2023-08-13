using System;

namespace WorkWithWord
{
    internal class JsonBody
    {
        [Newtonsoft.Json.JsonProperty("fileSavePath")]
        public string FileNamePath { get; set; } = Environment.CurrentDirectory;
    }
}
