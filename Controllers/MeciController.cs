using Microsoft.AspNetCore.Mvc;

namespace PricingService.Controllers
{
    [ApiController]

    [Route("[controller]")]
    public class MeciController : Controller
    {
        [HttpGet]
        public string Get()
        {
            var query = Request.QueryString.ToString();
            //return "Hello World";
            return ExcelCalculator.CalCulateMeci(query);
        }
    }
}
