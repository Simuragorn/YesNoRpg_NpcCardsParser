using CardsExcelParser.Enums;

namespace CardsExcelParser.Dtos
{
    public class NpcCardDto
    {
        public string NpcName { get; set; }
        public NpcEncounterTypeEnum NpcEncounterType { get; set; }
        public string DialogueText { get; set; }
        public List<EncounterResponseOption> ResponseOptions { get; set; }

        public NpcCardDto()
        {
            ResponseOptions = new List<EncounterResponseOption>();
        }
    }
}
