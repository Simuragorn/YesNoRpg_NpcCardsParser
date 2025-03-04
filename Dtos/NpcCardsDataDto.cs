
namespace CardsExcelParser.Dtos
{
    public class NpcCardsDataDto
    {
        public List<NpcCardConfigurationDto> NpcCardConfigurations;

        public NpcCardsDataDto()
        {
            NpcCardConfigurations = new List<NpcCardConfigurationDto>();
        }
    }
}
