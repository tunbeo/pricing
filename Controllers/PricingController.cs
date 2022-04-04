using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace PricingService.Controllers
{
    [ApiController]
    [Route("[controller]")]
   
    public class PricingController : ControllerBase
    {       
        [HttpGet]
        public Models.PriceResponse PricingExcel([FromQuery] string request)
        {
            return ExcelCalculator.CalculatePrice(request);
        }
    }
}
