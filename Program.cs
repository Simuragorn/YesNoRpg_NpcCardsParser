using CardsExcelParser.Constants;
using CardsExcelParser.Dtos;
using CardsExcelParser.Enums;
using CardsExcelParser.Extensions;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.ComponentModel;

namespace CardsExcelParser
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Enter the full path of the Excel file with npc cards: ");
            string excelFilePath = Console.ReadLine()?.Trim().Trim('"');

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                Console.WriteLine("Invalid file path. Please provide a valid Excel file.");
                return;
            }

            string npcCardsJsonOutputPath = Path.Combine(Path.GetDirectoryName(excelFilePath), "npc_cards.json");
            string multilanguageTextsJsonOutputPath = Path.Combine(Path.GetDirectoryName(excelFilePath), "multilanguage_texts.json");

            try
            {
                NpcCardsDataDto npcCardsData;
                MultilanguageTextDataDto multilanguageTextData;
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    npcCardsData = ParseNpcCardsFromExcelPackage(package);
                    multilanguageTextData = ParseMultilanguageTextsFromExcelPackage(package);
                }
                string npcCardsJson = JsonConvert.SerializeObject(npcCardsData, Formatting.Indented);
                string multilanguageTextJson = JsonConvert.SerializeObject(multilanguageTextData, Formatting.Indented);

                File.WriteAllText(npcCardsJsonOutputPath, npcCardsJson);
                File.WriteAllText(multilanguageTextsJsonOutputPath, multilanguageTextJson);

                Console.WriteLine("Excel file successfully converted to JSONs.");
                Console.WriteLine($"Output saved to: {npcCardsJsonOutputPath} and {multilanguageTextsJsonOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static MultilanguageTextDataDto ParseMultilanguageTextsFromExcelPackage(ExcelPackage package)
        {
            var result = new MultilanguageTextDataDto();
            ExcelWorksheet multilanguageTextsWorksheet = GetWorksheet(package, WorksheetConstants.MultilanguageTextsWorksheetName);
            int rowCount = multilanguageTextsWorksheet.Dimension.Rows;
            int colCount = multilanguageTextsWorksheet.Dimension.Columns;

            Dictionary<string, int> headers = GetHeaders(multilanguageTextsWorksheet, colCount);

            for (int row = 2; row <= rowCount; row++)
            {
                var textData = new TextDataDto();
                textData.Key = GetCellValue(multilanguageTextsWorksheet, row, headers, MultilanguageTextsHeaderColumns.KeyColumnName);

                if (string.IsNullOrEmpty(textData.Key))
                {
                    continue;
                }

                textData.Language = GetCellValue(multilanguageTextsWorksheet, row, headers, MultilanguageTextsHeaderColumns.LanguageColumnName);
                textData.Value = GetCellValue(multilanguageTextsWorksheet, row, headers, MultilanguageTextsHeaderColumns.ValueColumnName);

                result.TextDatas.Add(textData);
            }
            return result;
        }

        public static NpcCardsDataDto ParseNpcCardsFromExcelPackage(ExcelPackage package)
        {
            var result = new NpcCardsDataDto();
            ExcelWorksheet npcCardsWorksheet = GetWorksheet(package, WorksheetConstants.NpcCardsWorksheetName);
            int rowCount = npcCardsWorksheet.Dimension.Rows;
            int colCount = npcCardsWorksheet.Dimension.Columns;

            Dictionary<string, int> headers = GetHeaders(npcCardsWorksheet, colCount);

            for (int row = 2; row <= rowCount; row++)
            {
                var card = new NpcCardConfigurationDto();
                card.NpcName = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.NpcNameColumnName);

                if (string.IsNullOrWhiteSpace(card.NpcName))
                {
                    continue;
                }

                card.NpcImage = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.NpcImageColumnName);
                string npcEncounterTypeString = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.EncounterTypeColumnName);
                card.NpcEncounterType = (NpcEncounterTypeEnum)EnumHelpers.GetValueByDisplay(typeof(NpcEncounterTypeEnum), npcEncounterTypeString);
                card.DialogueText = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.DialogueColumnName);

                EncounterResponseOption affirmativeResponse = GetAffirmativeResponseOption(npcCardsWorksheet, headers, row);
                card.ResponseOptions.Add(affirmativeResponse);
                EncounterResponseOption negativeResponse = GetNegativeResponseOption(npcCardsWorksheet, headers, row);
                card.ResponseOptions.Add(negativeResponse);

                result.NpcCardConfigurations.Add(card);
            }
            return result;

            static EncounterResponseOption GetAffirmativeResponseOption(ExcelWorksheet worksheet, Dictionary<string, int> headers, int row)
            {
                string responseText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.AffirmativeResponseTextColumnName);
                string goldDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.GoldAffirmativeResponseColumnName);
                string materialsDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.MaterialsAffirmativeResponseColumnName);
                string reputationDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.ReputationAffirmativeResponseColumnName);
                EncounterResponseOption affirmativeResponse = ParseEncounterResponseOption(NpcResponseOptionTypeEnum.Affirmative, responseText, goldDeltaText, materialsDeltaText, reputationDeltaText);
                return affirmativeResponse;
            }
        }

        private static ExcelWorksheet GetWorksheet(ExcelPackage package, string worksheetName)
        {
            ExcelWorksheet npcCardsWorksheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == worksheetName);
            if (npcCardsWorksheet == null)
            {
                throw new Exception($"There is not worksheet with name: {worksheetName}");
            }

            return npcCardsWorksheet;
        }

        private static EncounterResponseOption GetNegativeResponseOption(ExcelWorksheet worksheet, Dictionary<string, int> headers, int row)
        {
            string responseText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.NegativeResponseTextColumnName);
            string goldDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.GoldNegativeResponseColumnName);
            string materialsDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.MaterialsNegativeResponseColumnName);
            string reputationDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.ReputationNegativeResponseColumnName);
            EncounterResponseOption negativeResponse = ParseEncounterResponseOption(NpcResponseOptionTypeEnum.Negative, responseText, goldDeltaText, materialsDeltaText, reputationDeltaText);
            return negativeResponse;
        }

        private static EncounterResponseOption ParseEncounterResponseOption(NpcResponseOptionTypeEnum type, string responseText, string goldDeltaText, string materialsDeltaText, string reputationDeltaText)
        {
            int.TryParse(goldDeltaText, out int goldDelta);
            int.TryParse(materialsDeltaText, out int materialsDelta);
            int.TryParse(reputationDeltaText, out int reputationDelta);

            var responseOption = new EncounterResponseOption
            {
                Type = type,
                ResponseText = responseText,
                GoldDelta = goldDelta,
                MaterialsDelta = materialsDelta,
                ReputationDelta = reputationDelta
            };
            return responseOption;
        }

        private static string GetCellValue(ExcelWorksheet worksheet, int row, Dictionary<string, int> headers, string columnName)
        {
            if (!headers.TryGetValue(columnName, out int columnIndex))
            {
                return string.Empty;
            }
            string cellValue = worksheet.Cells[row, columnIndex].Text?.Trim() ?? string.Empty;
            return cellValue;
        }

        private static Dictionary<string, int> GetHeaders(ExcelWorksheet worksheet, int colCount)
        {
            var headers = new Dictionary<string, int>();
            for (int col = 1; col <= colCount; col++)
            {
                string header = worksheet.Cells[1, col].Text.Trim();
                if (!string.IsNullOrEmpty(header))
                {
                    headers[header] = col;
                }
            }
            return headers;
        }
    }
}