using System;
using System.Collections.Generic;

namespace GostCode
{
    /// <summary>
    /// Group data
    /// </summary>
    class Group
    {
        /// <summary>
        /// User-friendly group name
        /// </summary>
        public String Name { get; set; }
        /// <summary>
        /// Group's annotation
        /// </summary>
        public String Annotation { get; set; }
        /// <summary>
        /// Filename filters (e.g. *.hpp)
        /// </summary>
        public List<String> Filters { get;  set; }
        /// <summary>
        /// Folders relative to base directory
        /// </summary>
        public List<String> Folders { get;  set; }
        /// <summary>
        /// Should the folder traversal be recursive?
        /// </summary>
        /// <value>true if recursive</value>
        /// <value>false if non-recursive</value>
        public bool Recursive { get; set; } = true;
    }
}