using Newtonsoft.Json;

namespace DLLCreateScheduleExcel.Entities
{
    public partial class Linha
    {
        [JsonProperty("tr")]
        public Tr[] Tr { get; set; }
    }
}