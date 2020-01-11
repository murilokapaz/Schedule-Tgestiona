using DLLCreateScheduleExcel.Services;
using Newtonsoft.Json;

namespace DLLCreateScheduleExcel.Entities
{
    public partial class Td
    {
        [JsonProperty("id")]
        [JsonConverter(typeof(ParseStringConverter))]
        public long Id { get; set; }

        [JsonProperty("valor")]
        public string Valor { get; set; }
    }
}
