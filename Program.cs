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
        private const string NpcNameColumnName = "NPC";
        private const string EncounterTypeColumnName = "Encounter Type";
        private const string DialogueColumnName = "Dialogue";

        private const string GoldAffirmativeResponseColumnName = "Gold (yes)";
        private const string MaterialsAffirmativeResponseColumnName = "Materials (yes)";
        private const string ReputationAffirmativeResponseColumnName = "Reputation (yes)";

        private const string GoldNegativeResponseColumnName = "Gold (no)";
        private const string MaterialsNegativeResponseColumnName = "Materials (no)";
        private const string ReputationNegativeResponseColumnName = "Reputation (no)";

        static void Main(string[] args)
        {
            Console.Write("Enter the full path of the Excel file: ");
            string excelFilePath = Console.ReadLine()?.Trim().Trim('"');

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                Console.WriteLine("Invalid file path. Please provide a valid Excel file.");
                return;
            }

            string jsonOutputPath = Path.Combine(Path.GetDirectoryName(excelFilePath), "npc_cards.json");

            try
            {
                var npcCardsData = ParseNpcCardsFromExcelToJson(excelFilePath);
                string json = JsonConvert.SerializeObject(npcCardsData, Formatting.Indented);
                File.WriteAllText(jsonOutputPath, json);

                Console.WriteLine("Excel file successfully converted to JSON.");
                Console.WriteLine($"Output saved to: {jsonOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static NpcCardsDataDto ParseNpcCardsFromExcelToJson(string filePath)
        {
            var result = new NpcCardsDataDto();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                Dictionary<string, int> headers = GetHeaders(worksheet, colCount);

                for (int row = 2; row <= rowCount; row++)
                {
                    var card = new NpcCardConfigurationDto();
                    card.NpcName = GetCellValue(worksheet, row, headers, NpcNameColumnName);
                    string npcEncounterTypeString = GetCellValue(worksheet, row, headers, EncounterTypeColumnName);
                    card.NpcEncounterType = (NpcEncounterTypeEnum)EnumHelpers.GetValueByDisplay(typeof(NpcEncounterTypeEnum), npcEncounterTypeString);
                    card.DialogueText = GetCellValue(worksheet, row, headers, DialogueColumnName);

                    EncounterResponseOption affirmativeResponse = GetAffirmativeResponseOption(worksheet, headers, row);
                    card.ResponseOptions.Add(affirmativeResponse);
                    EncounterResponseOption negativeResponse = GetNegativeResponseOption(worksheet, headers, row);
                    card.ResponseOptions.Add(negativeResponse);

                    result.NpcCardConfigurations.Add(card);
                }
            }
            return result;

            static EncounterResponseOption GetAffirmativeResponseOption(ExcelWorksheet worksheet, Dictionary<string, int> headers, int row)
            {
                string goldDeltaText = GetCellValue(worksheet, row, headers, GoldAffirmativeResponseColumnName);
                string materialsDeltaText = GetCellValue(worksheet, row, headers, MaterialsAffirmativeResponseColumnName);
                string reputationDeltaText = GetCellValue(worksheet, row, headers, ReputationAffirmativeResponseColumnName);
                EncounterResponseOption affirmativeResponse = ParseEncounterResponseOption(NpcResponseOptionTypeEnum.Affirmative, goldDeltaText, materialsDeltaText, reputationDeltaText);
                return affirmativeResponse;
            }
        }

        private static EncounterResponseOption GetNegativeResponseOption(ExcelWorksheet worksheet, Dictionary<string, int> headers, int row)
        {
            string goldDeltaText = GetCellValue(worksheet, row, headers, GoldNegativeResponseColumnName);
            string materialsDeltaText = GetCellValue(worksheet, row, headers, MaterialsNegativeResponseColumnName);
            string reputationDeltaText = GetCellValue(worksheet, row, headers, ReputationNegativeResponseColumnName);
            EncounterResponseOption negativeResponse = ParseEncounterResponseOption(NpcResponseOptionTypeEnum.Negative, goldDeltaText, materialsDeltaText, reputationDeltaText);
            return negativeResponse;
        }

        private static EncounterResponseOption ParseEncounterResponseOption(NpcResponseOptionTypeEnum type, string goldDeltaText, string materialsDeltaText, string reputationDeltaText)
        {
            int.TryParse(goldDeltaText, out int goldDelta);
            int.TryParse(materialsDeltaText, out int materialsDelta);
            int.TryParse(reputationDeltaText, out int reputationDelta);

            var responseOption = new EncounterResponseOption
            {
                Type = type,
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

        private static int GetIntValue(string value)
        {
            return int.TryParse(value, out int result) ? result : 0;
        }
    }
}