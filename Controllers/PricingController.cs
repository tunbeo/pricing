using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace PricingService.Controllers
{
    [Route("[controller]")]
    [ApiController]
    [Produces("application/json")]
    public class PricingController : ControllerBase
    {
        //[HttpGet]
        //public string Get()
        //{
        //    var query = Request.QueryString.ToString();
        //    return ExcelCalculator.Calculate(query);
        //}

        [HttpGet]
        public string Get()
        {
            var query = Request.QueryString.ToString();
            return JsonConvert.SerializeObject(ExcelCalculator.CalculatePrice(query));
        }
    }
}
