using CardsExcelParser.Enums;

namespace CardsExcelParser.Dtos
{
    public class NpcCardConfigurationDto
    {
        public List<MultilanguageTextDto> NpcNames;
        public string NpcImage;
        public NpcEncounterTypeEnum NpcEncounterType;
        public List<EncounterResponseOption> ResponseOptions;

        public List<MultilanguageTextDto> DialogueTexts;

        public NpcCardConfigurationDto()
        {
            ResponseOptions = new List<EncounterResponseOption>();
        }
    }
}
