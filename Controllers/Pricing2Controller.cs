using Microsoft.AspNetCore.Mvc;

namespace PricingService.Controllers
{
    public class Pricing2Controller : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public Models.PriceResponse PricingExcel([FromQuery] string request)
        {
            return ExcelCalculator.CalculatePrice(request);
        }
    }
}
