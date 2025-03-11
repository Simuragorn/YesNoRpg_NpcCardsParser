namespace CardsExcelParser.Dtos
{
    public class MultilanguageTextDataDto
    {
        public List<TextDataDto> TextDatas { get; set; }

        public MultilanguageTextDataDto()
        {
            TextDatas = new List<TextDataDto>();
        }
    }
}
