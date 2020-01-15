using DLLCreateScheduleExcel.Services;
using Newtonsoft.Json;

namespace DLLCreateScheduleExcel.Entities
{
    public partial class Coluna
    {
        [JsonProperty("id")]
        [JsonConverter(typeof(ParseStringConverter))]
        public long Id { get; set; }

        [JsonProperty("valor")]
        public string Valor { get; set; }

        [JsonProperty("elemento")]
        public string Elemento { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }
    }
}