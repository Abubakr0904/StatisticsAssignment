using Microsoft.AspNetCore.Mvc;
using StatisticsAssignment.Db;

namespace StatisticsAssignment.Controllers
{
    [Route("api/country")]
    [ApiController]
    public class CountryController : ControllerBase
    {
        [HttpGet("data")]
        public async Task<byte[]> GetData([FromServices] AssignmentDbContext dbContext)
        {
            return await new ExcelService(dbContext).GetExcelFile();
        }
    }
}
