using NPOI.POIFS.FileSystem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace OfficeFileProperties.Support.Npoi
{
    public class NpoiPropertiesOnlyDocument : NPOI.POIDocument
    {
        /// <summary>
        /// Create new document for accessing properties.
        /// </summary>
        /// <param name="fs">File system</param>
        public NpoiPropertiesOnlyDocument(NPOIFSFileSystem fs) : base(fs.Root)
        { }

        /// <summary>
        /// Create new document for accessing properties.
        /// </summary>
        /// <param name="fs">File system</param>
        public NpoiPropertiesOnlyDocument(POIFSFileSystem fs) : base(fs)
        { }

        /// <summary>
        /// Write to output stream. Not implemented.
        /// </summary>
        /// <param name="out1"></param>
        public override void Write(Stream out1)
        {
            throw new NotImplementedException();
        }
    }
}
