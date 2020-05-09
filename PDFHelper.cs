using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Collections;
using System.Drawing;
using System.Drawing.Imaging;

using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;

using O2S.Components.PDFRender4NET;

using Newtonsoft.Json;

namespace pdftool
{
     public class PDFHelper
    {
         public static void FillTemplete(string templetefile, string outputfile, string json)
         {


             Dictionary<string, string> data_dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);




             PdfReader pdfReader = new PdfReader(templetefile);
             PdfStamper pdfStamper = null;
             //try
             //{
             if (File.Exists(outputfile)) //若存在则退出
             {
                 return;
             }
             pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfile, FileMode.Create));
             //设置字体解决中文问题
             BaseFont font = BaseFont.CreateFont(@"C:\windows\fonts\simsun.ttc,1", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
             pdfStamper.AcroFields.AddSubstitutionFont(font);

             AcroFields pdfFormFields = pdfStamper.AcroFields;

             // set form pdfFormFields
             // 2015-8-23:以前循环的是json中的items,现在改为循环pdf中的fields
             //for (int x = 0; x < instObj.Items.Count; x++)
             //{
             //string fieldKey = instObj.Items[x].FieldKey;
             foreach (DictionaryEntry de in pdfReader.AcroFields.Fields)
             {
                 string fieldKey = de.Key.ToString();
                 System.Console.WriteLine("fieldKey:"+fieldKey);

                 ////直接读取
                 //string fieldValue = pdfJsonData.GetItem(fieldKey).FeildValue;// instObj.Items[x].FeildValue;


                 ////参考SrcField和RegEx读取
                 //PDFTempletFieldItem fi = this._fieldsInfo.GetItem(fieldKey);
                 //string srcField = Functions.ParseStr(fi.SrcField);
                 //if (!srcField.Equals(""))
                 //{
                 //    //srcfield的value
                 //    string srcFieldValue = Functions.ParseStr(pdfJsonData.GetItem(srcField).FeildValue);
                 //    string regex = Functions.ParseStr(fi.RegEx);
                 //    if (!srcFieldValue.Equals("") && !regex.Equals(""))
                 //    {
                 //        Regex r = new Regex(regex);
                 //        MatchCollection myMatches = r.Matches(srcFieldValue);
                 //        foreach (Match myMatch in myMatches)
                 //        {
                 //            fieldValue = myMatch.Groups[1].ToString();
                 //        }
                 //    }

                 //}
                 string fieldValue = "";
                 if (data_dict.ContainsKey(fieldKey))
                 {
                     fieldValue = data_dict[fieldKey];
                 }


                 pdfFormFields.SetField(fieldKey, fieldValue);
                 //pdfFormFields.SetFieldProperty(fieldKey, "setflags", PdfAnnotation.FLAGS_READONLY, null);

             }


             pdfStamper.FormFlattening = true; 
             // close the pdf
             pdfStamper.Close();

             //}
             //catch (Exception e)
             //{
             //    e.ToString();
             //    pdfStamper.Close();
             //    throw e;
             //}

         }


         /// <summary>
         /// 将PDF文档转换为图片的方法
         /// </summary>
         /// <param name="pdfInputPath">PDF文件路径</param>
         /// <param name="imageOutputPath">图片输出路径</param>
         /// <param name="imageName">生成图片的名字</param>
         /// <param name="startPageNum">从PDF文档的第几页开始转换</param>
         /// <param name="endPageNum">从PDF文档的第几页开始停止转换</param>
         /// <param name="imageFormat">设置所需图片格式</param>
         /// <param name="definition">设置图片的清晰度，数字越大越清晰</param>
         public static void ConvertPDF2Image(string pdfInputPath, string imageOutputPath,
             string imageName, int startPageNum, int endPageNum, ImageFormat imageFormat, int definition) //Definition definition)
         {
             PDFFile pdfFile = PDFFile.Open(pdfInputPath);

             if (!Directory.Exists(imageOutputPath))
             {
                 Directory.CreateDirectory(imageOutputPath);
             }

             // validate pageNum
             if (startPageNum <= 0)
             {
                 startPageNum = 1;
             }

             if (endPageNum > pdfFile.PageCount)
             {
                 endPageNum = pdfFile.PageCount;
             }

             if (startPageNum > endPageNum)
             {
                 int tempPageNum = startPageNum;
                 startPageNum = endPageNum;
                 endPageNum = startPageNum;
             }

             // start to convert each page
             for (int i = startPageNum; i <= endPageNum; i++)
             {
                 Bitmap pageImage = pdfFile.GetPageImage(i - 1, 56 * (int)definition);
                 string savefile = imageOutputPath +"\\"+ imageName +"_"+ i.ToString().PadLeft(3,'0') + "." + imageFormat.ToString();
                 System.Console.WriteLine("savefile:" + savefile);
                 if ( ! File.Exists(savefile)) {
                    pageImage.Save(savefile, imageFormat);
                    //pageImage.Save(imageOutputPath + imageName , imageFormat);
                 }
                 pageImage.Dispose();
             }

             pdfFile.Dispose();
         }


         public static void ConvertPDF2Png(string pdffilename,string outputpath)
         {
             System.Console.WriteLine("ConvertPDF2Png");
             PdfReader reader = new PdfReader(pdffilename);
             int numberOfPages = reader.NumberOfPages;
             System.Console.WriteLine("numberOfPages:" + numberOfPages);
             reader.Close();
             //string filepath = Path.GetDirectoryName(pdffilename) + "\\"; ;
             string imageName = Path.GetFileName(pdffilename);
             imageName = imageName.Substring(0, imageName.Length - 4);
             System.Console.WriteLine("outputpath:" + outputpath);
             System.Console.WriteLine("imageName:" + imageName);

             ConvertPDF2Image(pdffilename, outputpath, imageName, 1, numberOfPages, ImageFormat.Png, 4);

             
         }
    }
}
