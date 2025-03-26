using CardsExcelParser.Enums;

namespace CardsExcelParser.Dtos
{
    public class NpcCardConfigurationDto
    {
        public string NpcId;
        public List<MultilanguageTextDto> NpcNames;
        public string NpcImage;
        public NpcEncounterTypeEnum NpcEncounterType;
        public int? AgreementsCountRequired;
        public int? DisagreementsCountRequired;
        public List<EncounterResponseOption> ResponseOptions;

        public List<MultilanguageTextDto> DialogueTexts;

        public NpcCardConfigurationDto()
        {
            ResponseOptions = new List<EncounterResponseOption>();
        }
    }
}
