using System.ComponentModel.DataAnnotations;

namespace CardsExcelParser.Enums
{
    public enum NpcEncounterTypeEnum
    {
        [Display(Name = "Repeatable")]
        Repeatable = 0,
        [Display(Name = "Unique")]
        Unique = 1,
    }
}
