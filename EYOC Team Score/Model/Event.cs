using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EYOC_Team_Score.Model
{
    public enum EventType
    {
        Individual,
        Relay
    }

    public class Event
    {
        public required int Id { get; set; }
        public required string Name { get; set; }
        public required DateOnly Date { get; set; }
        public required EventType Type { get; set; }
        public List<Clazz> Clazzes { get; set; } = new List<Clazz>();

        public override string ToString()
        {
            return $"{Name} ({Date.ToString("o", CultureInfo.InvariantCulture)}), {Type}";
        }
    }

}
