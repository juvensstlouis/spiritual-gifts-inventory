using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using SpiritualGiftsInventory.Web.Excel;
using System.Linq;

namespace SpiritualGiftsInventory.Web.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class InventoryController : ControllerBase
    {
        [HttpPost]
        [Route("generate-formatted-excel")]
        public IActionResult GenerateFormattedExcel(IFormFile file)
        {
            if (file == null)
            {
                return BadRequest("O arquivo não foi informado.");
            }

            if (!file.FileName.EndsWith(".xlsx"))
            {
                return BadRequest("O arquivo informado é invalido.");
            }

            var records = ExcelReader.GetRecords(file);
            var fileContents = ExcelWriter.GenerateFormattedExcel(records.Where(r => r.Answers.Count() == 55));

            return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Dons Espirituais.xlsx");
        }
    }
}
