using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML;
using ClosedXML.Excel;
using SpiritualGiftsInventory.Web.Models;

public static class ExcelWriter
{
    private static int _currentRow;
    private static int _numberOfQuestionsPerGift = 5;

    public static byte[] GenerateFormattedExcel(IEnumerable<Member> records)
    {
        using var workbook = new XLWorkbook();

        foreach (var member in records)
        {
            string name = member.Name;

            if (member.Name.Length > 31)
            {
                name = member.Name.Substring(0, 30);
            }

            var worksheet = workbook.AddWorksheet(name);

            AddHeader(worksheet);
            AddGifts(worksheet, member.Answers);
            AddMemberInformation(worksheet, member);

            for (int i = 2; i <= 22; i += 2)
            {
                worksheet.Range(string.Format("F{0}:F{1}", i, i + 1)).Merge();
                worksheet.Range(string.Format("G{0}:G{1}", i, i + 1)).Merge();
            }

            worksheet.Range("A1:G23").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Range("A1:G23").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

            worksheet.Range("A1:G23").Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            worksheet.Range("A1:G23").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            worksheet.Range("A1:G23").Style.Fill.SetBackgroundColor(XLColor.AshGrey);

            worksheet.Range("I1:J4").Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            worksheet.Range("I1:J4").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            worksheet.Range("I1:J4").Style.Fill.SetBackgroundColor(XLColor.AshGrey);

            worksheet.ColumnsUsed().AdjustToContents();

            foreach (var column in worksheet.ColumnsUsed().Where(c => c.Width < 12))
            {
                column.Width = 12;
            };

            worksheet.PageSetup.SetPageOrientation(XLPageOrientation.Landscape);
            worksheet.PageSetup.SetPaperSize(XLPaperSize.A3Paper);
            worksheet.PageSetup.SetCenterHorizontally();
            worksheet.PageSetup.SetCenterVertically();

        }

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static void AddHeader(IXLWorksheet worksheet)
    {
        _currentRow = 1;

        for (int column = 1; column <= _numberOfQuestionsPerGift; column++)
        {
            worksheet.Cell(_currentRow, column).SetValue("*");
        }

        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 1).SetValue("TOTAL");
        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 2).SetValue("DONS");

        _currentRow++;
    }

    private static void AddGifts(IXLWorksheet worksheet, string[] answers)
    {
        AddGift(worksheet, answers, "Liderança", 1, 12, 23, 34, 45);
        AddGift(worksheet, answers, "Ex. Pessoal", 2, 13, 24, 35, 46);
        AddGift(worksheet, answers, "Organização", 3, 14, 25, 36, 47);
        AddGift(worksheet, answers, "Serviço", 4, 15, 26, 37, 48);
        AddGift(worksheet, answers, "Hospitalidade", 5, 16, 27, 38, 49);
        AddGift(worksheet, answers, "Intercessão", 6, 17, 28, 39, 50);
        AddGift(worksheet, answers, "Apostulado", 7, 18, 29, 40, 51);
        AddGift(worksheet, answers, "Assistência", 8, 19, 30, 41, 52);
        AddGift(worksheet, answers, "Cura e Saúde", 9, 20, 31, 42, 53);
        AddGift(worksheet, answers, "Pastoral", 10, 21, 32, 43, 54);
        AddGift(worksheet, answers, "Ensino", 11, 22, 33, 44, 55);
    }

    private static void AddGift(IXLWorksheet worksheet, string[] answers, string giftName, params int[] questionsNumbers)
    {
        int sum = 0;

        for (int i = 0; i < _numberOfQuestionsPerGift; i++)
        {
            int questionNumber = questionsNumbers[i];
            int answer = int.Parse(answers[questionNumber - 1]);
            sum += answer;

            worksheet.Cell(_currentRow, i + 1).SetValue("Pergunta " + questionNumber);
            worksheet.Cell(_currentRow + 1, i + 1).SetValue(answer);
        }
        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 1).SetValue(sum);
        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 2).SetValue(giftName);

        _currentRow += 2;
    }

    private static void AddMemberInformation(IXLWorksheet worksheet, Member member)
    {
        _currentRow = 1;

        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 4).SetValue("Nome");
        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 5).SetValue(member.Name);
        _currentRow++;

        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 4).SetValue("Email");
        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 5).SetValue(member.Email);
        _currentRow++;

        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 4).SetValue("Igreja");
        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 5).SetValue(member.Church);
        _currentRow++;

        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 4).SetValue("Data de Envio");
        worksheet.Cell(_currentRow, _numberOfQuestionsPerGift + 5).SetValue(member.SendDate);
        _currentRow++;
    }
}