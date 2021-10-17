using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using SpiritualGiftsInventory.Web.Models;

namespace SpiritualGiftsInventory.Web.Excel
{
    public static class ExcelReader
    {
        public static IEnumerable<Member> GetRecords(IFormFile file)
        {
            var records = new List<Member>();

            using var stream = new MemoryStream();
            file.CopyTo(stream);

            using var workbook = new XLWorkbook(stream);

            var firstWorkSheet = workbook.Worksheet(1);
            var rowsUsed = firstWorkSheet.RowsUsed().Skip(1);

            foreach (var row in rowsUsed)
            {
                var record = new Member
                {
                    SendDate = row.Cell(1)
                                  .GetDateTime()
                                  .ToShortDateString(),

                    Email = row.Cell(2).GetString().Trim(),
                    Punctuation = row.Cell(3).GetString().Trim(),
                    Name = row.Cell(4).GetString().Trim(),
                    Church = row.Cell(5).GetString().Trim(),

                    Answers = row.CellsUsed()
                                 .Skip(5)
                                 .Take(55)
                                 .Select(c => c.GetString().Trim())
                                 .ToArray()
                };

                records.Add(record);
            }

            return records;
        }
    }
}