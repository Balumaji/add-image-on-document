using DX4J.Image;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DX4J.FI.Console
{
    class Program
    {
        static void Main(string[] args)
        {            
            DX4JImage dx = new DX4JImage(@"C:\temp\docx4j_test.docx");
            dx.AddImageAnchor(@"C:\temp\bm.JPG",
                "Relationship of the Parties", 4, 3);

            dx.AddImageInline(@"C:\temp\bm.JPG",
                "Relationship of the Parties");
        }
    }
}
