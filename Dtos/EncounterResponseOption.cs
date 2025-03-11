using CardsExcelParser.Enums;

namespace CardsExcelParser.Dtos
{
    public class EncounterResponseOption
    {
        public NpcResponseOptionTypeEnum Type;
        public string ResponseTextEnglish;
        public string ResponseTextFrench;
        public int GoldDelta;
        public int MaterialsDelta;
        public int ReputationDelta;
    }
}
