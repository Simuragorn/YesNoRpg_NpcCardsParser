using CardsExcelParser.Constants;
using CardsExcelParser.Dtos;
using CardsExcelParser.Enums;
using CardsExcelParser.Extensions;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace CardsExcelParser
{
    internal class Program
    {
        private static ConsoleColor defaultConsoleColor;
        static void Main(string[] args)
        {
            defaultConsoleColor = Console.ForegroundColor;
            Console.Write("Enter the full path of the Excel file with npc cards: ");
            string excelFilePath = Console.ReadLine()?.Trim().Trim('"');

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                PrintError("Invalid file path. Please provide a valid Excel file.");
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

                PrintSuccess("Excel file successfully converted to JSONs.");
                PrintSuccess($"Output saved to: {npcCardsJsonOutputPath} and {multilanguageTextsJsonOutputPath}");
            }
            catch (Exception ex)
            {
                PrintError($"Error: {ex.Message}");
            }
        }

        private static void PrintSuccess(string printingText)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(printingText);
            Console.ForegroundColor = defaultConsoleColor;
        }
         private static void PrintError(string printingText)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(printingText);
            Console.ForegroundColor = defaultConsoleColor;
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
                card.EncounterId = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.EncounterIdColumnName);
                if (string.IsNullOrWhiteSpace(card.EncounterId))
                {
                    continue;
                }
                card.NpcId = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.NpcIdColumnName);
                if (int.TryParse(GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.ReputationResponseDeltaColumnName), out int reputationDelta))
                {
                    card.ReputationResponseDelta = reputationDelta;
                }
                if (int.TryParse(GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.AgreementsCountRequiredColumnName), out int agreementsCountRequired))
                {
                    card.AgreementsCountRequired = agreementsCountRequired;
                }
                if (int.TryParse(GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.DisagreementsCountRequiredColumnName), out int disagreementsCountRequired))
                {
                    card.DisagreementsCountRequired = disagreementsCountRequired;
                }
                List<string> npcNameHeaders = headers.Where(h => h.Key.Contains(NpcCardsHeaderColumns.NpcNamePartColumnName.Replace(" ", ""))).Select(h => h.Key).ToList();

                List<MultilanguageTextDto> npcNameTexts = new(npcNameHeaders.Count);
                foreach (var npcNameHeader in npcNameHeaders)
                {
                    string language = ExtractLanguageFromHeader(npcNameHeader);
                    string text = GetCellValue(npcCardsWorksheet, row, headers, npcNameHeader);
                    npcNameTexts.Add(new MultilanguageTextDto { Language = language, Text = text });
                }
                card.NpcNames = npcNameTexts;

                card.NpcImageRoot = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.NpcImageColumnName);
                card.MiniGameLevel = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.MiniGameLevelColumnName);
                string npcEncounterTypeString = GetCellValue(npcCardsWorksheet, row, headers, NpcCardsHeaderColumns.EncounterTypeColumnName);
                card.NpcEncounterType = (NpcEncounterTypeEnum)EnumHelpers.GetValueByDisplay(typeof(NpcEncounterTypeEnum), npcEncounterTypeString);
                List<string> dialogueHeaders = headers.Where(h => h.Key.Contains(NpcCardsHeaderColumns.DialoguePartColumnName.Replace(" ", ""))).Select(h => h.Key).ToList();
                List<MultilanguageTextDto> dialogueTexts = new(dialogueHeaders.Count);
                foreach (var dialogueHeader in dialogueHeaders)
                {
                    string language = ExtractLanguageFromHeader(dialogueHeader);
                    string text = GetCellValue(npcCardsWorksheet, row, headers, dialogueHeader);
                    dialogueTexts.Add(new MultilanguageTextDto { Language = language, Text = text });
                }
                card.DialogueTexts = dialogueTexts;

                EncounterResponseOption affirmativeResponse = GetAffirmativeResponseOption(npcCardsWorksheet, headers, row);
                card.ResponseOptions.Add(affirmativeResponse);
                EncounterResponseOption negativeResponse = GetNegativeResponseOption(npcCardsWorksheet, headers, row);
                card.ResponseOptions.Add(negativeResponse);

                result.NpcCardConfigurations.Add(card);
            }
            return result;

            static EncounterResponseOption GetAffirmativeResponseOption(ExcelWorksheet worksheet, Dictionary<string, int> headers, int row)
            {
                List<string> affirmativeResponseHeaders = headers.Where(h => h.Key.Contains(NpcCardsHeaderColumns.AffirmativeResponseTextPartColumnName.Replace(" ", ""))).Select(h => h.Key).ToList();
                List<MultilanguageTextDto> affirmativeResponseTexts = new(affirmativeResponseHeaders.Count);
                foreach (var affirmativeResponseHeader in affirmativeResponseHeaders)
                {
                    string language = ExtractLanguageFromHeader(affirmativeResponseHeader);
                    string text = GetCellValue(worksheet, row, headers, affirmativeResponseHeader);
                    affirmativeResponseTexts.Add(new MultilanguageTextDto { Language = language, Text = text });
                }

                string goldDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.GoldAffirmativeResponseColumnName);
                string materialsDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.MaterialsAffirmativeResponseColumnName);
                string reputationDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.ReputationAffirmativeResponseColumnName);

                EncounterResponseOption affirmativeResponse = ParseEncounterResponseOption(NpcResponseOptionTypeEnum.Affirmative, affirmativeResponseTexts, goldDeltaText, materialsDeltaText, reputationDeltaText);
                return affirmativeResponse;
            }
        }

        private static string ExtractLanguageFromHeader(string header)
        {
            int start = header.IndexOf('(');
            int end = header.IndexOf(')');

            if (start != -1 && end != -1 && end > start)
            {
                return header.Substring(start + 1, end - start - 1);
            }

            return header;
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
            List<string> negativeResponseHeaders = headers.Where(h => h.Key.Contains(NpcCardsHeaderColumns.NegativeResponseTextPartColumnName.Replace(" ", ""))).Select(h => h.Key).ToList();
            List<MultilanguageTextDto> negativeResponseTexts = new(negativeResponseHeaders.Count);
            foreach (var negativeResponseHeader in negativeResponseHeaders)
            {
                string language = ExtractLanguageFromHeader(negativeResponseHeader);
                string text = GetCellValue(worksheet, row, headers, negativeResponseHeader);
                negativeResponseTexts.Add(new MultilanguageTextDto { Language = language, Text = text });
            }
            string goldDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.GoldNegativeResponseColumnName);
            string materialsDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.MaterialsNegativeResponseColumnName);
            string reputationDeltaText = GetCellValue(worksheet, row, headers, NpcCardsHeaderColumns.ReputationNegativeResponseColumnName);
            EncounterResponseOption negativeResponse = ParseEncounterResponseOption(NpcResponseOptionTypeEnum.Negative, negativeResponseTexts, goldDeltaText, materialsDeltaText, reputationDeltaText);
            return negativeResponse;
        }

        private static EncounterResponseOption ParseEncounterResponseOption(NpcResponseOptionTypeEnum type, List<MultilanguageTextDto> responseTexts, string goldDeltaText, string materialsDeltaText, string reputationDeltaText)
        {
            int.TryParse(goldDeltaText, out int goldDelta);
            int.TryParse(materialsDeltaText, out int materialsDelta);
            int.TryParse(reputationDeltaText, out int reputationDelta);

            var responseOption = new EncounterResponseOption
            {
                Type = type,
                ResponseTexts = responseTexts,
                GoldDelta = goldDelta,
                MaterialsDelta = materialsDelta,
                ReputationDelta = reputationDelta
            };
            return responseOption;
        }

        private static string GetCellValue(ExcelWorksheet worksheet, int row, Dictionary<string, int> headers, string columnName)
        {
            if (!headers.TryGetValue(columnName.Replace(" ", ""), out int columnIndex))
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
                string header = worksheet.Cells[1, col].Text.Replace(" ", "");
                if (!string.IsNullOrEmpty(header))
                {
                    headers[header] = col;
                }
            }
            return headers;
        }
    }
}