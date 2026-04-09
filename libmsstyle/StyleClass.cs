using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace libmsstyle
{
    public class StyleClass
    {
        public int ClassId { get; set; }
        public string ClassName { get; set; }

        /// <summary>
        /// Base class ID from BCMAP. -1 means no explicit base class.
        /// </summary>
        public int BaseClassId { get; set; }

        public Dictionary<int, StylePart> Parts { get; set; }

        public StyleClass()
        {
            BaseClassId = -1;
            Parts = new Dictionary<int, StylePart>();
        }
    }
}
