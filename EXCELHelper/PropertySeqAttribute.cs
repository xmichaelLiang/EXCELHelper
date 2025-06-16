using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EXCELHelper {
   [AttributeUsage(AttributeTargets.Property)]
    public class PropertySeqAttribute: Attribute
    {
            public int Seq { get; }
            public PropertySeqAttribute(int seq)
            {
                 Seq = seq;
            }
    }
}