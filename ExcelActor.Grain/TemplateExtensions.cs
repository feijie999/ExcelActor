using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelActor.Grain
{
    public static class TemplateExtensions
    {
        public static Dictionary<string, object> AddNotNull(this Dictionary<string, object> source, string key, object value)
        {
            if (value != null)
                source.Add(key, value);

            return source;
        }

        public static Dictionary<string, object> AddNotNullOrSpace(this Dictionary<string, object> source, string key, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                source.Add(key, value);
            }

            return source;
        }

        public static Dictionary<string, object> AddNotNullOrSpace(this Dictionary<string, object> source, string key, Dictionary<string, object> value)
        {
            if (value != null && value.Count > 0)
            {
                source.Add(key, value);
            }

            return source;
        }

        public static Dictionary<string, object> UpdateNotNullOrSpace(this Dictionary<string, object> source, string key, Dictionary<string, object> value)
        {
            if (value != null && value.Count > 0)
            {
                if (!source.ContainsKey(key))
                    source.Add(key, value);
                else
                    source[key] = value;
            }

            return source;
        }

        public static Dictionary<string, object> UpdateNotNullOrSpace(this Dictionary<string, object> source, string key, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                if (!source.ContainsKey(key))
                    source.Add(key, value);
                else
                    source[key] = value;
            }

            return source;
        }

        public static Dictionary<string, object> Delete(this Dictionary<string, object> source, string key)
        {
            if (source.ContainsKey(key))
            {
                source.Remove(key);
            }

            return source;
        }

        public static Dictionary<string, object> Delete(this Dictionary<string, object> source, string key, int value)
        {
            if (source.ContainsKey(key))
            {
                if ((source[key] as int?) == value)
                    source.Remove(key);
            }

            return source;
        }

        public static string PixelWidth(this double width, decimal maxFontWidth = 7)
        {
            // 减去内空4px(左右各2px)
            return ((int)Math.Ceiling((double)maxFontWidth * width) - 4) + "px";
        }

        public static string PixelHeight(this double height)
        {
            return (int)(height * 96 / 72) + "px";
        }

        public static string PixelSize(this float size)
        {
            return (int)(size * 96 / 72) + "px";
        }
    }
}
