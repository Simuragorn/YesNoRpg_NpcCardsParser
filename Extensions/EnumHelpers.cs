using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace CardsExcelParser.Extensions
{
    public static class EnumHelpers
    {
        public static string GetDisplayValue(this Enum value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            var fieldInfo = value.GetType().GetField(value.ToString());
            if (fieldInfo == null)
            {
                return string.Empty;
            }

            DisplayAttribute displayAttribute = fieldInfo.GetCustomAttribute(
                typeof(DisplayAttribute), false) as DisplayAttribute;

            if (displayAttribute == null)
            {
                return String.Empty;
            }

            if (displayAttribute.ResourceType != null)
                return LookupResource(displayAttribute.ResourceType, displayAttribute.Name);

            return displayAttribute.Name;
        }

        public static Enum GetValueByDisplay(Type enumType, string display)
        {
            var values = enumType.GetEnumValues();
            foreach (Enum value in values)
            {
                if (value.GetDisplayValue() == display)
                {
                    return value;
                }
            }
            return null;
        }

        public static List<string> GetDisplayValues(Type enumType)
        {
            var displayValues = new List<string>();
            var values = enumType.GetEnumValues();
            foreach (Enum value in values)
            {
                displayValues.Add(value.GetDisplayValue());
            }
            return displayValues;
        }

        private static string LookupResource(Type resourceManagerProvider, string resourceKey)
        {
            foreach (PropertyInfo staticProperty in resourceManagerProvider.GetProperties(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public))
            {
                if (staticProperty.PropertyType == typeof(System.Resources.ResourceManager))
                {
                    System.Resources.ResourceManager resourceManager = (System.Resources.ResourceManager)staticProperty.GetValue(null, null);
                    return resourceManager.GetString(resourceKey);
                }
            }

            return resourceKey; // Fallback with the key name
        }
    }
}
