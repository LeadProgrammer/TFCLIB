using System;
using System.Collections.Generic;
using System.Text;

namespace LIBDI
{
    /// <summary>
    /// Exception class for CSVReader exceptions.
    /// </summary>
    public class CSVReaderException : ApplicationException
    {

        /// <summary>
        /// Constructs a new exception object with the given message.
        /// </summary>
        /// <param name="message">The exception message.</param>
        public CSVReaderException(string message) : base(message) { }
    }
}
