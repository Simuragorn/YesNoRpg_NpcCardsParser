using CardsExcelParser.Enums;

namespace CardsExcelParser.Dtos
{
    public class EncounterResponseOption
    {
        public NpcResponseOptionTypeEnum Type;
        public List<MultilanguageTextDto> ResponseTexts;
        public List<MultilanguageTextDto> AuthorTexts;
        public int GoldDelta;
        public int MaterialsDelta;
        public int ReputationDelta;
    }
}
