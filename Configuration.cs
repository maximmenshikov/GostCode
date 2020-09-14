using System;
using System.Collections.Generic;

namespace GostCode
{
    /// <summary>
    /// Global configuration object
    /// </summary>
    class Configuration
    {
        /// <summary>
        /// Style for group headers
        /// </summary>
        public String GroupStyle { get; set; }
        /// <summary>
        /// Style for group annotations
        /// </summary>
        public String AnnotationStyle { get; set; }
        /// <summary>
        /// Style for file name (section headers)
        /// </summary>
        /// <value></value>
        public String FileNameStyle { get; set; }
        /// <summary>
        /// Style for code block (section body)
        /// </summary>
        /// <value></value>
        public String CodeStyle { get; set; }
        /// <summary>
        /// Groups to be presented in the Word document
        /// </summary>
        /// <value></value>
        public List<Group> Groups { get; set; }
        /// <summary>
        /// Target project base directory. Can be empty, in that case,
        /// the one from parameters will be used.
        /// </summary>
        /// <value></value>
        public String BaseDirectory { get; set; } = "";
    }
}