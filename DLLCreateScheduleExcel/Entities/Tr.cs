using System;
using Newtonsoft.Json;

namespace DLLCreateScheduleExcel.Entities
{
    public partial class Tr
    {
        [JsonProperty("td")]
        public Td[] Td { get; set; }
    }
}
