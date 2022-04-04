using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace PricingService.Controllers
{
    [ApiController]
    [Produces("application/json")]
    [Route("[controller]")]
    public class PricingController : ControllerBase
    { 
        [HttpGet]
        public string Get()
        {
            var query = Request.QueryString.ToString();
            return ExcelCalculator.CalculatePrice(query);
        }
    }
}
