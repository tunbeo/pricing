using Microsoft.AspNetCore.Mvc;

namespace PricingService.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class PricingController : ControllerBase
    {
        [HttpGet]
        public string Get()
        {
            var query = Request.QueryString.ToString();
            return ExcelCalculator.Calculate(query);
        }
    }
}
