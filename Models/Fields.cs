//-----------------------------------------------------
// <copyright file="Fields.cs" company="Lateral Thinking Solutions Inc.">
//  Copyright 2016 James N. Romach, All rights Reserved
// </copyright>
// <summary>Object to store field data</summary>
//-----------------------------------------------------


namespace TfsDataReporting.Models
{
    public class Fields
    {
        /// <summary>
        /// Initializes a new instance of the Fields class
        /// </summary>
        /// <param name="name">String of the name</param>
        public Fields(string name)
        {
            this.FieldName = name;
        }

        /// <summary>
        /// Gets the string to hold field name
        /// </summary>
        public string FieldName { get; private set; }
    }
}