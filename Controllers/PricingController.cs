using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace PricingService.Controllers
{
    [ApiController]
    [Route("[controller]")]
   
    public class PricingController : ControllerBase
    { 
        [HttpGet]
        public string Get()
        {
            var query = Request.QueryString.ToString();
            return JsonConvert.SerializeObject(ExcelCalculator.CalculatePrice(query), Newtonsoft.Json.Formatting.None,
                            new JsonSerializerSettings
                            {
                                NullValueHandling = NullValueHandling.Ignore
                            });
        }
    }
}
