using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;

namespace libmsstyle.Import
{
    /// <summary>
    /// Small reflection based JSON writer for the import reports. Public readable properties,
    /// dictionaries, enumerables and primitives. No dependency on a json package.
    /// </summary>
    public static class JsonWriter
    {
        public static string Serialize(object value)
        {
            var sb = new StringBuilder();
            Write(sb, value, 0);
            sb.AppendLine();
            return sb.ToString();
        }

        private static void Write(StringBuilder sb, object v, int indent)
        {
            switch (v)
            {
                case null:
                    sb.Append("null");
                    return;
                case string s:
                    WriteString(sb, s);
                    return;
                case bool b:
                    sb.Append(b ? "true" : "false");
                    return;
                case Enum e:
                    WriteString(sb, e.ToString());
                    return;
                case int _:
                case long _:
                case short _:
                case ushort _:
                case uint _:
                case byte _:
                    sb.Append(Convert.ToString(v, CultureInfo.InvariantCulture));
                    return;
                case double d:
                    sb.Append(d.ToString("R", CultureInfo.InvariantCulture));
                    return;
                case IDictionary dict:
                    {
                        if (dict.Count == 0) { sb.Append("{}"); return; }
                        sb.Append("{\n");
                        bool first = true;
                        foreach (DictionaryEntry kv in dict)
                        {
                            if (!first) sb.Append(",\n");
                            first = false;
                            Indent(sb, indent + 1);
                            WriteString(sb, Convert.ToString(kv.Key, CultureInfo.InvariantCulture));
                            sb.Append(": ");
                            Write(sb, kv.Value, indent + 1);
                        }
                        sb.Append('\n');
                        Indent(sb, indent);
                        sb.Append('}');
                        return;
                    }
                case IEnumerable list:
                    {
                        var items = list.Cast<object>().ToList();
                        if (items.Count == 0) { sb.Append("[]"); return; }
                        sb.Append("[\n");
                        for (int i = 0; i < items.Count; ++i)
                        {
                            if (i > 0) sb.Append(",\n");
                            Indent(sb, indent + 1);
                            Write(sb, items[i], indent + 1);
                        }
                        sb.Append('\n');
                        Indent(sb, indent);
                        sb.Append(']');
                        return;
                    }
            }

            var props = v.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.GetIndexParameters().Length == 0 && p.GetCustomAttribute<JsonIgnoreAttribute>() == null)
                .ToList();
            if (props.Count == 0) { sb.Append("{}"); return; }
            sb.Append("{\n");
            for (int i = 0; i < props.Count; ++i)
            {
                if (i > 0) sb.Append(",\n");
                Indent(sb, indent + 1);
                WriteString(sb, props[i].Name);
                sb.Append(": ");
                Write(sb, props[i].GetValue(v), indent + 1);
            }
            sb.Append('\n');
            Indent(sb, indent);
            sb.Append('}');
        }

        private static void Indent(StringBuilder sb, int n)
        {
            sb.Append(' ', n * 2);
        }

        private static void WriteString(StringBuilder sb, string s)
        {
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("x4"));
                        else sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
        }
    }

    [AttributeUsage(AttributeTargets.Property)]
    public sealed class JsonIgnoreAttribute : Attribute { }
}
