using CardsExcelParser.Enums;

namespace CardsExcelParser.Dtos
{
    public class NpcCardConfigurationDto
    {
        public string NpcName;
        public string NpcImage;
        public NpcEncounterTypeEnum NpcEncounterType;
        public List<EncounterResponseOption> ResponseOptions;

        public string DialogueTextEnglish;
        public string DialogueTextFrench;

        public NpcCardConfigurationDto()
        {
            ResponseOptions = new List<EncounterResponseOption>();
        }
    }
}
