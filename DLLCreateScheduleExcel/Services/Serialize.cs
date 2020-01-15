using DLLCreateScheduleExcel.Entities;
using Newtonsoft.Json;

namespace DLLCreateScheduleExcel.Services
{
    public static class Serialize
    {
        public static string ToJson(this Welcome[] self) => JsonConvert.SerializeObject(self, Converter.Settings);
    }
}