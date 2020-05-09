using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;
using System.Text.RegularExpressions;



namespace pdftool
{
    class Program
    {
        static void Main(string[] args)
        {
            string cmd = "";


            cmd = args[0].Trim();
            if (cmd.ToLower().Equals("--fill")){
                string templetefile = args[1].Trim();
                string outputfile = args[2].Trim();
                string json = args[3].Trim();

                System.Console.WriteLine("templetefile:" + templetefile);
                System.Console.WriteLine("outputfile:" + outputfile);
                System.Console.WriteLine("json:" + json);

                PDFHelper.FillTemplete(templetefile, outputfile, json);

            }

            if (cmd.ToLower().Equals("--pdf2png")) {
                string inputfile = args[1].Trim();
                string outputpath = args[2].Trim();

                System.Console.WriteLine("inputfile:" + inputfile);
                System.Console.WriteLine("outputpath:" + outputpath);

                PDFHelper.ConvertPDF2Png(inputfile, outputpath);
            }

            //System.Console.Read();
            

        }
    }
}
