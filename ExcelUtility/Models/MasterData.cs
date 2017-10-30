using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUtility.Models
{
    public class MasterData
    {
        /// <summary>
        /// Gets or sets the identifier.
        /// </summary>
        /// <value>
        /// The identifier.
        /// </value>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the additional value.
        /// </summary>
        /// <value>
        /// The additional value.
        /// </value>
        public string AdditionalValue { get; set; }

        /// <summary>
        /// Gets or sets the additional value2.
        /// </summary>
        /// <value>
        /// The additional value2.
        /// </value>
        public string AdditionalValue2 { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is service.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance is service; otherwise, <c>false</c>.
        /// </value>
        public bool IsService { get; set; }
    }
}