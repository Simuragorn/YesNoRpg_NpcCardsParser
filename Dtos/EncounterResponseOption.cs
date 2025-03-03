using CardsExcelParser.Enums;

namespace CardsExcelParser.Dtos
{
    public class EncounterResponseOption
    {
        public NpcResponseOptionTypeEnum Type { get; set; }
        public int GoldDelta { get; set; }
        public int MaterialsDelta { get; set; }
        public int ReputationDelta { get; set; }
    }
}
