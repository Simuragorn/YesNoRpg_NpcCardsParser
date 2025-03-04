using CardsExcelParser.Enums;

namespace CardsExcelParser.Dtos
{
    public class NpcCardConfigurationDto
    {
        public string NpcName;
        public NpcEncounterTypeEnum NpcEncounterType;
        public string DialogueText;
        public List<EncounterResponseOption> ResponseOptions;

        public NpcCardConfigurationDto()
        {
            ResponseOptions = new List<EncounterResponseOption>();
        }
    }
}
