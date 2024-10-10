/* 
  Program: Casenamics Document Agent

  Copyright (C) 2015, Casenamics - All Rights Reserved
  Unauthorized copying of this file, via any medium is strictly prohibited
  Proprietary and confidential
  Written by Brandon Path

  Description:

  Takes an input file(pdf or tiff), splits it to files of one page length, reads the barcode/s in each page and 
  merges the pages that have matching barcode/s into a single file and passes the file/barcode as argument to a vbscript


  Files are copied to the Folder_Input. Once they are processed, the script will move them to Folder_Input_Processed. 

  TODO: Support BMP, JPG, PNG, etc. 

  // PDFSharp files are in:
  // C:\Workbnch\CaseZoo\DocumentConverter\packages\PDFsharp.1.32.3057.0\lib\net20

 */



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;

using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Security;

using System.Collections;
using System.Configuration;
using System.Diagnostics;
using System.Threading;
using System.Xml;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using ZXing;
using IronOcr;
using Tesseract;


using Leadtools;
using Leadtools.Codecs;



using Google.Cloud.DocumentAI.V1;
using Google.Protobuf;
using System.Text.Json;
using Grpc.Core;




// TODO: optimize, if no split page, then just copy file instead of creating new tif faster?
// TODO: (DONE) How to process files on initial service start instead of having to copy a new file to it. 
// TODO: (FUTURE) Convert PDF to TIFF before it can decode barcode.. 
// TODO: (FUTURE) Read the PDF file spec, use open source PDF analyzer, then extract each image, and save it as TIFF format. 
// TODO: Then, can use Zxing to read barcode values..
// Then, disrupt the market with this product!!!
//
//   Bring scanner home to stress test with different barcodes and directions... 
// TODO: Sign Files. 
// TODO: Document how to setup and installation instructions.. in readme.txt.  already created. 

// NOTE: 
//
// PDFSharp, only supports PDF Version 5.0 and lower. 



namespace Casenamics
{

   
    /// <summary>
    /// 
    /// </summary>
    public class DocumentAgent
    {

        enum LICENSE_VERSION { INVALID, FREE, TRIAL, STANDARD, PRO};
        enum FILE_FORMAT { PDF, TIFF, BMP, GIF, JPG, PNG };
        int threadcount = 0;
        public  string Folder_Input;
        public string Folder_Input_Temp;
        public string Folder_Input_Processed; // just a sub folder of Input folder.. 
        public  string Folder_Output_Success;
        public  string Folder_Input_Work;
        public  string Folder_Output_Error;
        public  string Folder_Scripts;
        // 9-19-2018, null check causing object ref. error, so set logFileName to blank.. 
        public  string LogFileName = "";
        public  int Number_Of_Threads;
        public string BarcodeFormats;
        public string LicenseFile;
        public string TrialFile;
        public int TimerInterval;
        public bool VerboseLogging = false;
        public string projectId = "204385916807";
        public string locationId = "us";
        public string processorId = "2bda2437616e1291";
        public string localPath = "C:\\junk\\Combined_Document.tiff";
        public string mimeType = "image/tiff";
        public string outputTextPath = "C:/output-text.txt";
        public string outputJsonPath = "C:/output-data.json";



        Queue file_queue = Queue.Synchronized(new Queue());
        private bool EnablePDF = true; //  enable or disable processing of PDF files.. 


        public bool IsTrialVersion(ref int daysLeft)
        {
            bool RetVal = false; daysLeft = 0;

            // returns whether the license is a trial or not..
            LICENSE_VERSION LicenseVersion = GetProductLicense(LicenseFile);
            if (LicenseVersion == LICENSE_VERSION.TRIAL)
            {
                daysLeft = CheckTrialDays(TrialFile);
                RetVal = true;
            }


            return (RetVal);
        }

        // start it up
        public void StartAgent()
        {
            bool validLicense = true;
            try
            {
                CreateFolders(); // create folders if they don't exist
                CreateConfigFileForScript(); // create config file for script to read from

                LICENSE_VERSION licenseVersion = GetProductLicense(LicenseFile);
                WriteLog($"License Type = {licenseVersion}");

                switch (licenseVersion)
                {
                    case LICENSE_VERSION.TRIAL:
                        WriteLog("License Detected = Trial");
                        if (CheckTrialDays(TrialFile) > 30)
                        {
                            WriteLog("Your 30-day trial license has expired. Please contact sales@casenamics.com for license registration.");
                            validLicense = false;
                        }
                        break;

                    case LICENSE_VERSION.INVALID:
                        WriteLog("Your license is invalid. Please contact sales@casenamics.com for license registration.");
                        validLicense = false;
                        break;

                    case LICENSE_VERSION.FREE:
                        WriteLog("License Detected = Free");
                        if (CheckHowManyFilesProcessedToday() > 25)
                        {
                            WriteLog("You have a free license. You are not allowed to process more than 25 documents per day. Please contact sales@casenamics.com for license registration.");
                            validLicense = false;
                        }
                        break;
                }

                if (validLicense)
                {
                    ProcessDocuments();
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }


        // checks if file is being used by another process..
        // doesn't work.. 
        //bool FileInUse(string filePath)
        //{
        //    bool RetVal = false;
        //    try
        //    {
        //        FileAttributes attributes = File.GetAttributes(filePath);
        //        RetVal = false;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(filePath + " was used by another process. Did not add to queue.");
        //        RetVal = true;
        //    }
        //    return (RetVal);
        //}

        //populate the file queue for processing
        void PopulateQueue(ArrayList items)
        {
            try
            {
                for (int i = 0; i < items.Count; i++)
                {
                    // 5-29-2018, bpath, only add file to the queue if you have access to it
                    // if it's in use, do not add it. Otherwise, you can not move it.. 
                    //if (!FileInUse(items[i].ToString()))
                    //{
                        file_queue.Enqueue(items[i]);
                        WriteLog(items[i] + " added to the file queue.");
                    //}

                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }

        }
        

        //Write configuration settings to log file.. 
        void PrintSettings()
        {
            try
            {
                WriteLog("Folder_Input:" + Folder_Input);
                WriteLog("Folder_Input_Temp:" + Folder_Input_Temp);
                WriteLog("Folder_Input_Processed:" + Folder_Input_Processed);
                WriteLog("Folder_Input_Work: " + Folder_Input_Work);
                WriteLog("Folder_Output_Success: " + Folder_Output_Success);
                WriteLog("Folder_Output_Error: " + Folder_Output_Error);
                WriteLog("Folder_Threads: " + threadcount.ToString());
                WriteLog("Timer Interval: " + TimerInterval.ToString());
                WriteLog("Verbose Logging: " + (VerboseLogging ? "true" : "false"));
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }



        /*Split Multipage Tiff File into individual pages and save them in the FOLDER_INPUT_WORK
         * 
         * 
         * @Params
         * inputPath-> FileName(includeing its full path)
         * outputPath-> The directory where the split files are stored
         * @return
         * pages-> returns the number of pages after split
         */

        ///
        protected int SplitTiffAndSave(string inputPath, string outputPath)
        {
            int ActivePage;
            int Pages=0;
            try
            {
                System.Drawing.Image Image = System.Drawing.Image.FromFile(inputPath);
                Pages = Image.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);
                for (int Index = 0; Index < Pages; Index++)
                {

                    ActivePage = Index + 1;
                    Image.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, Index);
                    Image.Save(outputPath + "\\" + ActivePage.ToString().PadLeft(10, '0') + ".tiff"); // padding zeros since we need to sort by filename.. 

                }
                Image.Dispose(); // dispose so it closes file.. 
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            return Pages;
        }

        ///*Split Multipage pdf File into individual pages and save them in the FOLDER_INPUT_WORK
        // * 
        // * 
        // * @Params
        // * inputPath-> FileName(includeing its full path)
        // * outputPath-> The directory where the split files are stored
        // * @return
        // * pages-> returns the number of pages after split
        // */
        int SplitAndSave(string inputPath, string outputPath)
        {

            PdfSharp.Pdf.PdfDocument inputDocument = new PdfSharp.Pdf.PdfDocument();  
            try
            {
                inputDocument = PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfDocumentOpenMode.Import);
                string name = Path.GetFileNameWithoutExtension(inputPath);

                for (int PageNumber = 0; PageNumber < inputDocument.PageCount; PageNumber++)
                {
                    // Create new document
                    PdfSharp.Pdf.PdfDocument outputDocument = new PdfSharp.Pdf.PdfDocument();
                    outputDocument.Version = inputDocument.Version;
                    outputDocument.Info.Title =
                      String.Format("Page {0} of {1}", PageNumber, inputDocument.Info.Title);
                    outputDocument.Info.Creator = inputDocument.Info.Creator;

                    string Filename = outputPath + @"\" + PageNumber.ToString().PadLeft(10, '0') + ".pdf";  // padding zeros since we need to sort by filename... 

                    // Add the page and save it
                    outputDocument.AddPage(inputDocument.Pages[PageNumber]);
                    outputDocument.Save(Filename);
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            return (inputDocument.PageCount);
        }


        // Create the File name that gets created in the Work folder.
        String CreateFileIDForZXing(ZXing.Result[] barcode, string inputFile)
        {

            String OutputFileName="";
            try
            {

                
                //string InputFileName = inputFile.Substring(Folder_Input.Length + 1).Split('.')[0];
                // get filename without extension.. 
                string InputFileName = Path.GetFileName(inputFile).Split('.')[0];
                StringBuilder Sb = new StringBuilder(); // "File_"
                Sb.Append(InputFileName);
                if (barcode != null)
                {
                    for (int i = 0; i < barcode.Length; i++)
                    {
                        Sb.Append("_"); // add dash
                        Sb.Append(barcode[i].Text);
                    }
                }

                else {

                    Sb.Append("_NoBarcode");

                }

                OutputFileName = Sb.ToString();
                WriteLog("Filename created  " + OutputFileName);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }

            return OutputFileName;

        }

        
        /*merge tiff files into a multi-page tiff file and save
         * 
         * @params
         * str_DestinationPath-> the output file path 
         * sourceFiles-> array of filenames of the files to be merged
         * 
         */
        public  void mergeTiffPages(string str_DestinationPath, string[] sourceFiles)
        {

            System.Drawing.Imaging.ImageCodecInfo codec = null;

            foreach (System.Drawing.Imaging.ImageCodecInfo cCodec in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
            {
                if (cCodec.CodecName == "Built-in TIFF Codec")
                    codec = cCodec;
            }

            try
            {

                System.Drawing.Imaging.EncoderParameters imagePararms = new System.Drawing.Imaging.EncoderParameters(1);
                imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.MultiFrame);

                if (sourceFiles.Length == 1)
                {
                    System.IO.File.Copy((string)sourceFiles[0], str_DestinationPath, true);

                }
                else if (sourceFiles.Length >= 1)
                {
                    System.Drawing.Image DestinationImage = (System.Drawing.Image)(new System.Drawing.Bitmap((string)sourceFiles[0]));

                    DestinationImage.Save(str_DestinationPath, codec, imagePararms);

                    imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.FrameDimensionPage);

                    //                    for (int i = 0; i < sourceFiles.Length - 1; i++)

                    for (int i = 1; i < sourceFiles.Length; i++)
                    {
                        System.Drawing.Image img = (System.Drawing.Image)(new System.Drawing.Bitmap((string)sourceFiles[i]));

                        DestinationImage.SaveAdd(img, imagePararms);
                        img.Dispose();
                    }

                    imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.Flush);
                    DestinationImage.SaveAdd(imagePararms);
                    imagePararms.Dispose();
                    DestinationImage.Dispose();

                }

            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        // merge the individual pages back into one and save it.
        void SaveOutputFile(string outputFilePath, FILE_FORMAT fileFormat, string[] inputFileNames, int startIndex, int endIndex)
        {
            try
            {
                if (fileFormat == FILE_FORMAT.PDF)
                {
                    PdfSharp.Pdf.PdfDocument OutputDocument = new PdfSharp.Pdf.PdfDocument();

                    for (int i = startIndex; i <= endIndex; i++)
                    {
                        PdfSharp.Pdf.PdfDocument InputDocument = PdfSharp.Pdf.IO.PdfReader.Open(inputFileNames[i], PdfDocumentOpenMode.Import);
                        OutputDocument.AddPage(InputDocument.Pages[0]);
                    }
                    OutputDocument.Save(outputFilePath);
                    OutputDocument.Close();
                }
                else
                {
                    string[] sourcefiles = new string[endIndex - startIndex + 1];
                    for (int i = startIndex; i <= endIndex; i++)
                    {
                        Array.Copy(inputFileNames, startIndex, sourcefiles, 0, (endIndex - startIndex + 1));
                    }
                    mergeTiffPages(outputFilePath, sourcefiles);
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }


        // decode the barcodes in the file older version
        //private ZXing.Result[] DecodeBarcodeZxing(string fileName, FILE_FORMAT fileFormat  =FILE_FORMAT.BMP)
        //{
        //    ZXing.Result[] result = new ZXing.Result[] { };
        //    ZXing.BarcodeReader reader = new ZXing.BarcodeReader();

        //    reader.Options.PossibleFormats = new List<ZXing.BarcodeFormat>();

        //    // EAN_8, EAN_13, ITF, the commandlineencoder demo project failed to generate these barcodes.
        //    // UPC-A, UPC-E, EAN-8, EAN-13, Code 39, Code 93, Code 128, ITF, Codabar, MSI, RSS-14, QR Code, Data Matrix, Aztec and PDF-417

        //    if (BarcodeFormats.ToUpper() == "ALL_1D")
        //    {
        //        reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.All_1D);
        //    }
        //    else
        //    {

        //        if (BarcodeFormats.IndexOf("CODE_39") >= 0)
        //        {
        //            reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.CODE_39);
        //            reader.Options.UseCode39ExtendedMode = true;
        //            //reader.Options.UseCode39RelaxedExtendedMode = true;

        //        }
        //        if (BarcodeFormats.IndexOf("CODE_93") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.CODE_93);
        //        if (BarcodeFormats.IndexOf("CODE_128") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.CODE_128);
        //        if (BarcodeFormats.IndexOf("QR_CODE") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.QR_CODE);
        //        if (BarcodeFormats.IndexOf("EAN_8") >= 0)  reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.EAN_8);
        //        if (BarcodeFormats.IndexOf("UPC_EAN_EXTENSION") >= 0)  reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.UPC_EAN_EXTENSION);
        //        if (BarcodeFormats.IndexOf("EAN_13") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.EAN_13);
        //        if (BarcodeFormats.IndexOf("UPC_A") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.UPC_A);
        //        if (BarcodeFormats.IndexOf("UPC_E") >= 0)  reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.UPC_E);
        //        if (BarcodeFormats.IndexOf("ITF") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.ITF);
        //        if (BarcodeFormats.IndexOf("PDF_417") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.PDF_417);
        //        if (BarcodeFormats.IndexOf("CODABAR") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.CODABAR);
        //        if (BarcodeFormats.IndexOf("MSI") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.MSI);
        //        if (BarcodeFormats.IndexOf("RSS-14") >= 0)  reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.RSS_14);
        //        if (BarcodeFormats.IndexOf("RSS_EXPANDED") >= 0)  reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.RSS_EXPANDED);
        //        if (BarcodeFormats.IndexOf("DATA_MATRIX") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.DATA_MATRIX);
        //        if (BarcodeFormats.IndexOf("AZTEC") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.AZTEC);
        //    }

        //    reader.Options.TryHarder = true; // process more instead of trying to go for speed. 
        //   // reader.Options.PureBarcode = true; // purebarcode is if image is strictly monochrome. 

        //    try
        //    {
        //        Bitmap BitmapFile;
        //        // bpath, 7-2-2018, if PDF, then need to extract image first.. 
        //        if (fileFormat == FILE_FORMAT.PDF)
        //        {
        //            BitmapFile = GetImageFromPDF(fileName);
        //            result = reader.DecodeMultiple(BitmapFile);
        //        }
        //        else // for PNG, TIFF, etc., this is supported.. 
        //        {
        //            Console.WriteLine("Decoding image: {0}", fileName);
        //            //Bitmap BitmapFile;
        //            BitmapFile = (Bitmap)Bitmap.FromFile(fileName);
        //            result = reader.DecodeMultiple(BitmapFile);
        //            //BitmapFile.Dispose(); // free up the file... 
        //        }
        //        BitmapFile.Dispose(); // free up the file... 
        //    }
        //    catch (Exception exc)
        //    {
        //        Console.WriteLine("Exception: {0}", exc.Message);
        //    }
        //    return (result);

        //}



        //**************************************************************************************************
        //Newer version of Barcode working with PDF's
        //**************************************************************************************************
        private ZXing.Result[] DecodeBarcodeZxing(string fileName, FILE_FORMAT fileFormat = FILE_FORMAT.BMP)
        {
            List<ZXing.Result> results = new List<ZXing.Result>();
            ZXing.BarcodeReader reader = new ZXing.BarcodeReader();

            reader.Options.PossibleFormats = new List<ZXing.BarcodeFormat>();

            // EAN_8, EAN_13, ITF, the commandlineencoder demo project failed to generate these barcodes.
            // UPC-A, UPC-E, EAN-8, EAN-13, Code 39, Code 93, Code 128, ITF, Codabar, MSI, RSS-14, QR Code, Data Matrix, Aztec and PDF-417

            if (BarcodeFormats.ToUpper() == "ALL_1D")
            {
                reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.All_1D);
            }
            else
            {
                if (BarcodeFormats.IndexOf("CODE_39") >= 0)
                {
                    reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.CODE_39);
                    reader.Options.UseCode39ExtendedMode = true;
                    //reader.Options.UseCode39RelaxedExtendedMode = true;
                }
                if (BarcodeFormats.IndexOf("CODE_93") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.CODE_93);
                if (BarcodeFormats.IndexOf("CODE_128") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.CODE_128);
                if (BarcodeFormats.IndexOf("QR_CODE") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.QR_CODE);
                if (BarcodeFormats.IndexOf("EAN_8") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.EAN_8);
                if (BarcodeFormats.IndexOf("UPC_EAN_EXTENSION") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.UPC_EAN_EXTENSION);
                if (BarcodeFormats.IndexOf("EAN_13") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.EAN_13);
                if (BarcodeFormats.IndexOf("UPC_A") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.UPC_A);
                if (BarcodeFormats.IndexOf("UPC_E") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.UPC_E);
                if (BarcodeFormats.IndexOf("ITF") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.ITF);
                if (BarcodeFormats.IndexOf("PDF_417") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.PDF_417);
                if (BarcodeFormats.IndexOf("CODABAR") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.CODABAR);
                if (BarcodeFormats.IndexOf("MSI") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.MSI);
                if (BarcodeFormats.IndexOf("RSS-14") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.RSS_14);
                if (BarcodeFormats.IndexOf("RSS_EXPANDED") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.RSS_EXPANDED);
                if (BarcodeFormats.IndexOf("DATA_MATRIX") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.DATA_MATRIX);
                if (BarcodeFormats.IndexOf("AZTEC") >= 0) reader.Options.PossibleFormats.Add(ZXing.BarcodeFormat.AZTEC);
            }

            reader.Options.TryHarder = true;

            ZXing.Result[] result = new ZXing.Result[] { };

            try
            {
                Bitmap BitmapFile;
                if (fileFormat == FILE_FORMAT.PDF)
                {
                    List<Bitmap> bitmaps = GetImagesFromPDF(fileName);
                    foreach (Bitmap bitmap in bitmaps)
                    {
                        result = reader.DecodeMultiple(bitmap);
                        if (result != null)
                        {
                            results.AddRange(result);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Decoding image: {0}", fileName);
                    BitmapFile = (Bitmap)Bitmap.FromFile(fileName);
                    result = reader.DecodeMultiple(BitmapFile);
                    if (result != null)
                    {
                        results.AddRange(result);
                    }
                    BitmapFile.Dispose();
                }
            }
            catch (Exception exc)
            {
                Console.WriteLine("Exception: {0}", exc.Message);
            }

            return results.ToArray();
        }






        //Updated Supporting function for scanning barcode from PDF.

        private List<Bitmap> GetImagesFromPDF(string fileName)
        {
            List<Bitmap> returnObjects = new List<Bitmap>();

            try
            {
                using (PdfDocument document = PdfReader.Open(fileName, PdfDocumentOpenMode.Import))
                {
                    foreach (PdfPage page in document.Pages)
                    {
                        foreach (Image image in page.GetImages())
                        {
                            if (image != null)
                            {
                                returnObjects.Add(new Bitmap(image));
                            }
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                Console.WriteLine("Exception: {0}", exc.Message);
            }

            return returnObjects;
        }









        // assumes one page file.. returns the first image.. 
        Bitmap GetImageFromPDF(string fileName)



        {


           // Bitmap ReturnObject = new Bitmap(1, 1); // just a bitmap size 1 x 1.. 

            //try
            //{
             //   using (PdfDocument document = PdfReader.Open(fileName, PdfDocumentOpenMode.Import))
              //  {
               //     int pageIndex = 0;
                //    foreach (PdfPage page in document.Pages)
                 //   {
                        // int imageIndex = 0;
                  //      foreach (Image image in page.GetImages())
                    //    {
                            //Console.WriteLine("\r\nExtracting image {1} from page {0}", pageIndex + 1, imageIndex + 1);
                            // Save the file images to disk in the current directory.
                            //image.Save(String.Format(@"{0:00000000}-{1:000}.png", pageIndex + 1, imageIndex + 1, Path.GetFileName(filename)), ImageFormat.Png);
                            //imageIndex++;
                      //      ReturnObject = new Bitmap(image);
                       // }
                       // pageIndex++;
                    //}
                //}
            //}
           Bitmap ReturnObject = null;

    try
    {
        using (PdfDocument document = PdfReader.Open(fileName, PdfDocumentOpenMode.Import))
        {
            int pageIndex = 0;
            foreach (PdfPage page in document.Pages)
            {
                foreach (Image image in page.GetImages())
                {
                    if (image != null)
                    {
                        ReturnObject = new Bitmap(image);
                        // You may want to break here if you only want the first image
                        break;
                    }
                }
                pageIndex++;
            }
        }
    }
            catch (Exception)
            {

                throw;
            }
            return (ReturnObject);
        }
        /*Takes an input file(pdf or tiff), splits it to files of one page length, reads the barcode/s in each page and 
         * merges the pages that have matching barcode/s into a single file and passes the file/barcode as argument to a vbscript
         * 
         * @params
         * inputFile-> the input filename(including the path)
         * fileFormat-> PDF or TIFF
         * workFolderPath-> the temporary folder where the split files are stored
         * ouputPath-> The folder where processed files are saved
         *  
         */
        void processFileUsingZXing(int threadID)
        {

            string InputFile;
            while (file_queue.Count > 0)
            {
                InputFile = file_queue.Dequeue().ToString();
                // Makes sure is not being used by another process before we process it.. 
                //if (!FileInUse(InputFile))
                //{
                    ProcessFile(InputFile, threadID);
                //}
            }
            WriteLog("FIle queue empty.. Exiting Thread" + threadID);
            return;

        }

        void ProcessFile(string inputFile, int threadID)
        {
            String WorkFolderPath;
            DirectoryInfo Dir; bool FileUsedByAnotherProcess = false;

            // OCR processing using IronOCR
            //var Ocr = new IronTesseract();
            //Ocr.Language = OcrLanguage.English;
            string extractedText = "";
   
            WriteLog("File started process \n");
            WriteLog("Processing File  " + inputFile + " from thread " + threadID);

            try
            {

                // Dir = Directory.CreateDirectory((Folder_Input + "//" + threadID));

                // bpath, 9-9-2018, changed to temp folder for the work folder.. 
                Dir = Directory.CreateDirectory((Folder_Input_Temp + "//" + threadID));
                WriteLog("My threadID is " + threadID);
                WorkFolderPath = Dir.FullName;

                //// move file to Temp folder so no other app can open it and prevent us from processing it..
                string TempFolder = Folder_Input_Temp + @"\" + threadID + @"\Temp";
                Directory.CreateDirectory(TempFolder);
                string DestInputFile = TempFolder + @"\" + Path.GetFileName(inputFile);
                try
                {
                    // remove destination file if it exists.. 
                    if (File.Exists(DestInputFile))
                    {
                        File.Delete(DestInputFile);
                    }
                    File.Move(inputFile, DestInputFile);
                }
                catch
                {
                    // if can't move the file, then the file is in use, so do not process it..
                    FileUsedByAnotherProcess = true;
                    WriteLog(inputFile + " is in use. Can not move it to the Temp Folder from thread " + threadID + ".");
                }

                if (!FileUsedByAnotherProcess)
                {
                    // now, set the inputFile to the new location. 
                    inputFile = DestInputFile;
                    String OutputPath = Folder_Input_Work;
                    /******************************************************************************************************/
                    /*  IF MULTIPLE PAGES, MUST SPLIT THE FILE SO WE CAN THEN SEPARATE THE ACTUAL DOCUMENTS FROM THE FILE */
                    /******************************************************************************************************/

                    // bpath, 7-22-2018, rewrite to get the file extension correctly.. If filename has period, this would
                    // cause an issue..
                    string FileExt = Path.GetExtension(inputFile);
                    FILE_FORMAT fileFormat;
                    if (FileExt.ToLower() == ".pdf")
                        fileFormat = FILE_FORMAT.PDF;
                    else if (FileExt.ToLower() == ".tif" || FileExt.ToLower() == ".tiff")
                        fileFormat = FILE_FORMAT.TIFF;
                    else if (FileExt.ToLower() == ".bmp")
                        fileFormat = FILE_FORMAT.BMP;
                    else if (FileExt.ToLower() == ".jpg")
                        fileFormat = FILE_FORMAT.JPG;
                    else if (FileExt.ToLower() == ".gif")
                        fileFormat = FILE_FORMAT.GIF;
                    else if (FileExt.ToLower() == ".png")
                        fileFormat = FILE_FORMAT.PNG;
                    else
                    {
                        fileFormat = new FILE_FORMAT();
                        WriteLog("FileFormat not recognized");
                    }



                    //string[] dummy = inputFile.Split('.');
                    //FILE_FORMAT fileFormat;
                    //if ((dummy[dummy.Length - 1]).ToLower() == "pdf")
                    //    fileFormat = FILE_FORMAT.PDF;
                    //else if ((dummy[dummy.Length - 1]).ToLower() == "tif" || (dummy[dummy.Length - 1]).ToLower() == "tiff")
                    //    fileFormat = FILE_FORMAT.TIFF;
                    //else if ((dummy[dummy.Length - 1]).ToLower() == "bmp")
                    //    fileFormat = FILE_FORMAT.BMP;
                    //else if ((dummy[dummy.Length - 1]).ToLower() == "jpg")
                    //    fileFormat = FILE_FORMAT.JPG;
                    //else if ((dummy[dummy.Length - 1]).ToLower() == "gif")
                    //    fileFormat = FILE_FORMAT.GIF;
                    //else if ((dummy[dummy.Length - 1]).ToLower() == "png")
                    //    fileFormat = FILE_FORMAT.PNG;
                    //else
                    //{
                    //    fileFormat = new FILE_FORMAT();
                    //    WriteLog("FileFormat not recognized");
                    //}
                    WriteLog("This is a " + fileFormat.ToString() + " file");
                    if (fileFormat == FILE_FORMAT.PDF)
                    {
                        //Split the pdf into individual pages
                        int numPagesSaved = SplitAndSave(inputFile, WorkFolderPath);
                    }
                    else if (fileFormat == FILE_FORMAT.TIFF)
                    {
                        int numPagesSaved = SplitTiffAndSave(inputFile, WorkFolderPath);
                    }
                    else
                    {
                        if (!(fileFormat == FILE_FORMAT.BMP || fileFormat == FILE_FORMAT.JPG || fileFormat == FILE_FORMAT.GIF || fileFormat == FILE_FORMAT.PNG))
                        {
                            WriteLog("Unsupported File Format");
                            return;
                        }
                    }

                    /**********************************************************************************************************************************************************************/
                    /*  Now that the file itself has been split into individual pages, lets search for the pages that belong to one file, and then merge those pages back into one file.  */
                    /*  If individual file is one document on its own, no merging is needed, but the file and its metadata will still be created. */
                    /**********************************************************************************************************************************************************************/
                    

                    
                    
                    
                    string[] SplitPagesFileList;
                    //Read bar codes in each page


                    if (fileFormat == FILE_FORMAT.PDF)
                    {
                        var ReturnList = Directory.EnumerateFiles(WorkFolderPath, "*.pdf", SearchOption.TopDirectoryOnly).OrderBy(filename => filename); // return sorted..
                        SplitPagesFileList = ReturnList.ToArray();
                    }
                    else if (fileFormat == FILE_FORMAT.TIFF)
                    {
                        var ReturnList = Directory.EnumerateFiles(WorkFolderPath, "*.tiff", SearchOption.TopDirectoryOnly).OrderBy(filename => filename); // return sorted.. 
                        SplitPagesFileList = ReturnList.ToArray();
                    }
                    // If format is BMP, PNG, GIF, JPG, then no need to split. Set first record in array to the filename itself... 
                    else if (fileFormat == FILE_FORMAT.BMP || fileFormat == FILE_FORMAT.JPG || fileFormat == FILE_FORMAT.GIF || fileFormat == FILE_FORMAT.PNG)
                    {
                        SplitPagesFileList = new string[1];
                        SplitPagesFileList[0] = inputFile;
                    }
                    else
                    {
                        WriteLog("Unsupported File Format during split");
                        return;
                    }

                    /**********************************************************************************************************************************************************************/
                    //Updated code of Google Document AI with KEY VaLUE OUTPUT AND DATA TABLE stored in DAXML and JSON
                    /**********************************************************************************************************************************************************************/

                    var keyValueString = ""; var tableString=""; var jsonoutput = "";
                        var client = new DocumentProcessorServiceClientBuilder
                        {
                            Endpoint = $"{locationId}-documentai.googleapis.com"
                        }.Build();

                    //try
                    //{

                    //    var fileStream = File.OpenRead(inputFile);
                    //    var rawDocument = new RawDocument
                    //    {
                    //        Content = ByteString.FromStream(fileStream),
                    //        MimeType = GetMimeType(inputFile) // Function to get the MIME type based on file extension
                    //    };

                    //    // Initialize request argument(s)
                    //    var request = new ProcessRequest
                    //    {
                    //        Name = ProcessorName.FromProjectLocationProcessor(projectId, locationId, processorId).ToString(),
                    //        RawDocument = rawDocument
                    //    };

                    //    // Make the request
                    //    var response = client.ProcessDocument(request);

                    //    var document = response.Document;
                    //    Console.WriteLine(document.Text);

                    //    // Optionally, handle the extracted data
                    //    (keyValueString, tableString, jsonoutput) = ExtractData(document, outputTextPath, outputJsonPath);

                    //    extractedText = keyValueString + tableString;
                    //    Console.WriteLine("Comp Data" + extractedText);
                    //    Console.WriteLine("JSON data " + jsonoutput);

                    //}
                    //catch (IOException ioEx)
                    //{
                    //    Console.WriteLine($"IOException: {ioEx.Message}");
                    //    // Handle file IO exceptions (e.g., file not found, file in use)
                    //}


                    // Read in local file
                     extractedText = "";
                    try
                    {
                        
                        // Ensure the file is not being used by another process
                        using (var fileStream = File.OpenRead(inputFile))
                        {
                            //sqs
                            var rawDocument = new RawDocument
                            {
                                Content = ByteString.FromStream(fileStream),
                                MimeType = GetMimeType(inputFile) // Function to get the MIME type based on file extension
                            };

                            // Initialize request argument(s)
                            var request = new ProcessRequest
                            {
                                Name = ProcessorName.FromProjectLocationProcessor(projectId, locationId, processorId).ToString(),
                                RawDocument = rawDocument
                            };

                            // Make the request
                            var response = client.ProcessDocument(request);

                            var document = response.Document;
                            Console.WriteLine(document.Text);

                            // Optionally, handle the extracted data
                             (keyValueString, tableString, jsonoutput) = ExtractData(document, outputTextPath, outputJsonPath);

                            extractedText = keyValueString + tableString;
                            Console.WriteLine(extractedText);
                            Console.WriteLine(jsonoutput);

                            string outputFilePath = Path.Combine(Path.GetDirectoryName(OutputPath)+"/Success", Path.GetFileNameWithoutExtension(inputFile) + ".json");

                            // Write the JSON output to the file
                            File.WriteAllText(outputFilePath, jsonoutput);
                            Console.WriteLine($"JSON output saved to {outputFilePath}");
                        }
                    }
                    catch (IOException ioEx)
                    {
                        Console.WriteLine($"IOException: {ioEx.Message}");
                        
                    }
                    catch (RpcException rpcEx)
                    {
                        Console.WriteLine($"RpcException: {rpcEx.Status}");
                        Console.WriteLine($"RpcException Message: {rpcEx.Message}");
                        Console.WriteLine($"RpcException Details: {rpcEx.Status.Detail}");
                     
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Exception: {ex.Message}");
                     
                    }






                    /**********************************************************************************************************************************************************************/

                    //using IRONOCR
                    //foreach (var pageFile in SplitPagesFileList)
                    //{
                    //    var OcrResult = Ocr.Read(pageFile);
                    //    if (OcrResult != null && OcrResult.Text != null)
                    //    {
                    //        // Extracted text from OCR
                    //        extractedText = OcrResult.Text;
                    //        WriteLog(extractedText);


                    //    }
                    //    else
                    //    {
                    //        // OCR failed or returned no text
                    //        WriteLog($"OCR failed for page: {pageFile}");
                    //    }
                    //}
                    /**********************************************************************************************************************************************************************/
                    /*  Merge files accordingly now so that each file has the same barcodes */
                    /**********************************************************************************************************************************************************************/
                    int StartIndex = 0; int EndIndex = 0;
                    string PreviousPageID = "", CurrentPageID = "";
                    var barcodes_start_page = DecodeBarcodeZxing(SplitPagesFileList[0].ToString(), fileFormat);

                    // get the first page and set the PreviousPageID .. 
                    PreviousPageID = CreateFileIDForZXing(barcodes_start_page, inputFile);
                    StartIndex = 0; EndIndex = 0;

                    // if only one page, save it... 
                    if (SplitPagesFileList.Length == 1)
                    {
                        MergeAndSaveFileForZXing(threadID, OutputPath, PreviousPageID, fileFormat, StartIndex, EndIndex, SplitPagesFileList, barcodes_start_page, inputFile, extractedText, jsonoutput);
                    }
                    else
                    {
                        // loop through remaining pages.. SplitPagesFileListCounter starts at 1, not 0 (since already processed above).. 
                        for (int SplitPagesFileListCounter = 1; SplitPagesFileListCounter < SplitPagesFileList.Length; SplitPagesFileListCounter++)
                        {
                            // Get barcodes for the current page... 
                            ZXing.Result[] barcodes_current_page = new ZXing.Result[] { };
                            barcodes_current_page = DecodeBarcodeZxing(SplitPagesFileList[SplitPagesFileListCounter].ToString(), fileFormat);

                            if (barcodes_current_page == null) CurrentPageID = PreviousPageID; // if no barcode on the current page, then it belongs to the previous page.. 
                            else CurrentPageID = CreateFileIDForZXing(barcodes_current_page, inputFile);

                            if (CurrentPageID != PreviousPageID)
                            {
                                // Found a new document, so save the previous document. 
                                MergeAndSaveFileForZXing(threadID, OutputPath, PreviousPageID, fileFormat, StartIndex, EndIndex, SplitPagesFileList, barcodes_start_page, inputFile, extractedText, jsonoutput);

                                // now, reset the indexes to begin new with the next document.. 
                                StartIndex = SplitPagesFileListCounter;
                                EndIndex = SplitPagesFileListCounter;
                                barcodes_start_page = barcodes_current_page; // set the new barcodes to the new page.. 
                            }
                            else
                            {
                                EndIndex = SplitPagesFileListCounter;
                            }
                            PreviousPageID = CurrentPageID; // now, set this as the previous page ID
                                                            // If this is the last page, we need to also save the output file... 
                            if (SplitPagesFileListCounter == SplitPagesFileList.Length - 1) MergeAndSaveFileForZXing(threadID, OutputPath, PreviousPageID, fileFormat, StartIndex, EndIndex, SplitPagesFileList, barcodes_start_page, inputFile, extractedText, jsonoutput);
                        } // for
                    }

                    if (fileFormat == FILE_FORMAT.PDF || fileFormat == FILE_FORMAT.TIFF)
                    {
                        /* Remove the split page files, only need to remove if files were PDF or TIFF since splitting occurs only with these formats */
                        foreach (string f in SplitPagesFileList)
                            File.Delete(f);
                    }

                    /* move original input file into processed folder, making copy if already exists. */
                    string OriginalDestFile = Folder_Input_Processed + @"\" + Path.GetFileName(inputFile);
                    string DestFile = OriginalDestFile; int FileCount = 1;
                    while (File.Exists(DestFile))
                    {
                        DestFile = Folder_Input_Processed + @"\" + Path.GetFileNameWithoutExtension(OriginalDestFile) + " (Copy " + FileCount + ")" + Path.GetExtension(OriginalDestFile);
                        FileCount++;
                    }
                    File.Move(inputFile, DestFile);
                }
            }
            catch (Exception e)
            {
                WriteLog("Exception: " + e.Message);
            }
        }

        private string GetMimeType(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            switch (extension)
            {
                case ".pdf":
                    return "application/pdf";
                case ".tiff":
                case ".tif":
                    return "image/tiff";
                case ".bmp":
                    return "image/bmp";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".png":
                    return "image/png";
                default:
                    throw new NotSupportedException("Unsupported file format");
            }
        }


        private static Document ProcessDocument(DocumentProcessorServiceClient client, string projectId, string locationId, string processorId, string filePath, string mimeType)
        {
            string resourceName = $"projects/{projectId}/locations/{locationId}/processors/{processorId}";

            ByteString content = ByteString.CopyFrom(File.ReadAllBytes(filePath));
            RawDocument rawDocument = new RawDocument
            {
                Content = content,
                MimeType = mimeType
            };

            ProcessRequest request = new ProcessRequest
            {
                Name = resourceName,
                RawDocument = rawDocument
            };

            ProcessResponse response = client.ProcessDocument(request);
            return response.Document;
        }

        private static (string keyValueString, string tableString, string jsonoutput) ExtractData(Document document, string outputTextPath, string outputJsonPath)
        {
            var keyValuePairs = new List<Dictionary<string, object>>();
            var tables = new List<Dictionary<string, object>>();

            List<string> names = new List<string>();
            List<float> nameConfidence = new List<float>();
            List<string> values = new List<string>();
            List<float> valueConfidence = new List<float>();

            foreach (var page in document.Pages)
            {
                if (page.FormFields != null)
                {
                    foreach (var field in page.FormFields)
                    {
                        var kvp = new Dictionary<string, object>();
                        if (field.FieldName?.TextAnchor?.Content != null)
                        {
                            string name = TrimText(field.FieldName.TextAnchor.Content);
                            name = name.Replace(":", ""); // Remove colon from the name

                            kvp["FieldName"] = name;
                            // kvp["FieldNameConfidence"] = field.FieldName.Confidence;
                            names.Add(name);
                            nameConfidence.Add(field.FieldName.Confidence);

                        }
                        else
                        {
                            kvp["FieldName"] = "null";
                            // kvp["FieldNameConfidence"] = 0;
                            names.Add("null");
                            nameConfidence.Add(0);
                        }

                        if (field.FieldValue?.TextAnchor?.Content != null)
                        {
                            string value = TrimText(field.FieldValue.TextAnchor.Content);
                            kvp["FieldValue"] = value;
                            // kvp["FieldValueConfidence"] = field.FieldValue.Confidence;
                            values.Add(value);
                            valueConfidence.Add(field.FieldValue.Confidence);
                        }
                        else
                        {
                            kvp["FieldValue"] = "null";
                            //  kvp["FieldValueConfidence"] = 0;
                            values.Add("null");
                            valueConfidence.Add(0);
                        }

                        keyValuePairs.Add(kvp);
                    }
                }

                // Extract table data
                if (page.Tables != null)
                {
                    foreach (var table in page.Tables)
                    {
                        var tableDict = new Dictionary<string, object>
                        {
                            ["PageNumber"] = page.PageNumber,
                            ["HeaderRows"] = new List<List<string>>(),
                            ["BodyRows"] = new List<List<string>>()
                        };

                        foreach (var row in table.HeaderRows)
                        {
                            var headerRow = new List<string>();
                            foreach (var cell in row.Cells)
                            {
                                headerRow.Add(TrimText(GetTextFromAnchor(cell.Layout.TextAnchor, document.Text)));
                            }
                            ((List<List<string>>)tableDict["HeaderRows"]).Add(headerRow);
                        }

                        foreach (var row in table.BodyRows)
                        {
                            var bodyRow = new List<string>();
                            foreach (var cell in row.Cells)
                            {
                                bodyRow.Add(TrimText(GetTextFromAnchor(cell.Layout.TextAnchor, document.Text)));
                            }
                            ((List<List<string>>)tableDict["BodyRows"]).Add(bodyRow);
                        }

                        tables.Add(tableDict);
                    }
                }
            }


            // Prepare the key-value string
            StringBuilder keyValueStringBuilder = new StringBuilder();
            keyValueStringBuilder.AppendLine("Key-Value Pairs:");
            for (int i = 0; i < names.Count; i++)
            {
                keyValueStringBuilder.AppendLine($"Field Name: {names[i]}");
                keyValueStringBuilder.AppendLine($"Field Value: {values[i]}");
                keyValueStringBuilder.AppendLine(new string('-', 50)); // Separator line
                keyValueStringBuilder.AppendLine();
            }
            string keyValueString = keyValueStringBuilder.ToString();

            // Prepare the table string
            StringBuilder tableStringBuilder = new StringBuilder();
            tableStringBuilder.AppendLine("Table Data:");
            foreach (var table in tables)
            {
                tableStringBuilder.AppendLine($"Table found on page {table["PageNumber"]}");
                tableStringBuilder.AppendLine(new string('-', 50)); // Separator line

                tableStringBuilder.AppendLine("Headers:");
                foreach (var row in (List<List<string>>)table["HeaderRows"])
                {
                    foreach (var cell in row)
                    {
                        tableStringBuilder.AppendLine($"  {cell}");
                    }
                }

                tableStringBuilder.AppendLine("\nRows:");
                foreach (var row in (List<List<string>>)table["BodyRows"])
                {
                    foreach (var cell in row)
                    {
                        tableStringBuilder.AppendLine($"  {cell}");
                    }
                }

                tableStringBuilder.AppendLine(new string('=', 50)); // End of table separator line
                tableStringBuilder.AppendLine("\n"); // Add a newline for spacing between tables
            }
            string tableString = tableStringBuilder.ToString();



            // Write the text output file
            //using (StreamWriter writer = new StreamWriter(outputTextPath))
            //{
            //    writer.WriteLine("Key-Value Pairs:");
            //    for (int i = 0; i < names.Count; i++)
            //    {
            //        writer.WriteLine($"Field Name: {names[i]}");
            //        writer.WriteLine($"Field Value: {values[i]}");
            //        writer.WriteLine(new string('-', 50)); // Separator line
            //        writer.WriteLine();
            //    }

            //    writer.WriteLine("Table Data:");
            //    foreach (var table in tables)
            //    {
            //        writer.WriteLine($"Table found on page {table["PageNumber"]}");
            //        writer.WriteLine(new string('-', 50)); // Separator line

            //        writer.WriteLine("Headers:");
            //        foreach (var row in (List<List<string>>)table["HeaderRows"])
            //        {
            //            foreach (var cell in row)
            //            {
            //                writer.WriteLine($"  {cell}");
            //            }
            //        }

            //        writer.WriteLine("\nRows:");
            //        foreach (var row in (List<List<string>>)table["BodyRows"])
            //        {
            //            foreach (var cell in row)
            //            {
            //                writer.WriteLine($"  {cell}");
            //            }
            //        }

            //        writer.WriteLine(new string('=', 50)); // End of table separator line
            //        writer.WriteLine("\n"); // Add a newline for spacing between tables
            //    }
            //}

            // Write the JSON output file
            var outputJson = new
            {
                KeyValuePairs = keyValuePairs,
                Tables = tables
            };

            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonString = JsonSerializer.Serialize(outputJson, jsonOptions);

            //File.WriteAllText(outputJsonPath, jsonString);

            //Console.WriteLine($"Output written to {outputTextPath} and {outputJsonPath}");

            return (keyValueString, tableString,jsonString);
        }

        private static string TrimText(string text)
        {
            return text?.Trim().Replace("\n", " ");
        }

        private static string GetTextFromAnchor(Document.Types.TextAnchor textAnchor, string text)
        {
            if (textAnchor == null || textAnchor.TextSegments == null)
            {
                return null;
            }

            string result = "";
            foreach (var segment in textAnchor.TextSegments)
            {
                int startIndex = (int)segment.StartIndex;
                int endIndex = (int)segment.EndIndex;
                result += text.Substring(startIndex, endIndex - startIndex);
            }

            return result;
        }




        // merge files back into one.
        void MergeAndSaveFileForZXing(int threadID, string outputPath, string pageID, FILE_FORMAT fileFormat, int startIndex, int endIndex, string[] SplitPagesFileList, ZXing.Result[] barcodes_page, string parentFileName, string extractedtext, string jsonoutput)
        {
            try
            {
                string OutputFileName;
                List<string> BarCodeList = new List<string>();

                for (int i = 0; i != barcodes_page.Count(); i++)
                {
                    BarCodeList.Add(barcodes_page[i].Text);
                }
                // file for DAXML
                OutputFileName = GenerateFileName(outputPath, pageID, GetFileFormatExtension(fileFormat));
                SaveOutputFile(OutputFileName, fileFormat, SplitPagesFileList, startIndex, endIndex);
                // create meta data file.. 
                CreateMetaDataFile(threadID, OutputFileName, BarCodeList, parentFileName,extractedtext, jsonoutput);

                //file for JSON

                string OutputFileNameJson = GenerateFileName(outputPath, pageID, ".json");
                SaveOutputFile(OutputFileNameJson, fileFormat, SplitPagesFileList, startIndex, endIndex);
                CreateMetaDataJsonFile(threadID, OutputFileNameJson, BarCodeList, parentFileName, jsonoutput);

                //string OutputFileNameJson = GenerateFileName(outputPath, pageID, GetFileFormatExtension(fileFormat));
                //SaveOutputFile(OutputFileNameJson, fileFormat, SplitPagesFileList, startIndex, endIndex);
                //CreateMetaDataJsonFile(threadID, OutputFileName, BarCodeList, parentFileName, jsonoutput);


            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }


        void CreateMetaDataJsonFile(int threadID, string outputFileName, List<string> barCodeList, string parentFileName, string jsonoutput)
        {
            try
            {
                var metadata = new
                {
                    ThreadID = threadID,
                    OutputFileName = outputFileName,
                    BarCodes = barCodeList,
                    ParentFileName = parentFileName,
                    JsonOutput = jsonoutput
                };
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(jsonoutput, jsonOptions);
                string metadataFileName = Path.ChangeExtension(outputFileName, ".json");
                File.WriteAllText(metadataFileName, json);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }



        string GetFileFormatExtension(FILE_FORMAT fileFormat)
        {
            string RetVal = "";
            switch (fileFormat)
            {
                case FILE_FORMAT.PDF:
                    RetVal = ".pdf";
                    break;
                case FILE_FORMAT.TIFF:
                    RetVal = ".tif";
                    break;
                case FILE_FORMAT.BMP:
                    RetVal = ".bmp";
                    break;
                case FILE_FORMAT.GIF:
                    RetVal = ".gif";
                    break;
                case FILE_FORMAT.JPG:
                    RetVal = ".jpg";
                    break;
                case FILE_FORMAT.PNG:
                    RetVal = ".png";
                    break;
            }
            return (RetVal);
        }

        string GenerateFileName(string outputPath, string ID, string fileExtension)
        {
            string OutputFilePath = outputPath + @"\" + ID + fileExtension;
            string OutputFileSearchDir = outputPath + @"\" + ID + "*" + fileExtension;

            try
            {
                // if file already exists, then append a copy to it... 
                if (File.Exists(OutputFilePath))
                {
                    // see how many are out there.. 
                    string[] FileNames = Directory.GetFiles(outputPath, ID + "*" + fileExtension, SearchOption.TopDirectoryOnly);
                    int Count = FileNames.Length;

                    OutputFilePath = outputPath + @"\" + ID + " (Copy " + Count + ")" + fileExtension;
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            return (OutputFilePath);
        }

        // Loop through all files in the folder and process each file. 
         void ProcessDocuments()
        {
            try
            {
                String inputPath = Folder_Input;

                ArrayList Input_Filenames = new ArrayList();

                if (EnablePDF)
                {
                    var Input_Filenames_Pdf = Directory.EnumerateFiles(inputPath, "*.pdf", SearchOption.TopDirectoryOnly).OrderBy(filename => filename);
                    foreach (string currentFile in Input_Filenames_Pdf)
                    {
                        Input_Filenames.Add(currentFile);
                    }
                }

                var Input_Filenames_Tif = Directory.EnumerateFiles(inputPath, "*.tiff", SearchOption.TopDirectoryOnly).OrderBy(filename => filename);
                foreach (string currentFile in Input_Filenames_Tif)
                {
                    Input_Filenames.Add(currentFile);
                }


                // Get BMP Files
                var Input_Filenames_BMP = Directory.EnumerateFiles(inputPath, "*.bmp", SearchOption.TopDirectoryOnly).OrderBy(filename => filename);
                foreach (string currentFile in Input_Filenames_BMP)
                {
                    Input_Filenames.Add(currentFile);
                }

                // Get GIF Files
                var Input_Filenames_GIF = Directory.EnumerateFiles(inputPath, "*.gif", SearchOption.TopDirectoryOnly).OrderBy(filename => filename);
                foreach (string currentFile in Input_Filenames_GIF)
                {
                    Input_Filenames.Add(currentFile);
                }

                // Get JPEG Files
                var Input_Filenames_JPEG = Directory.EnumerateFiles(inputPath, "*.jpg", SearchOption.TopDirectoryOnly).OrderBy(filename => filename);
                foreach (string currentFile in Input_Filenames_JPEG)
                {
                    Input_Filenames.Add(currentFile);
                }

                // Get PNG Files
                var Input_Filenames_PNG = Directory.EnumerateFiles(inputPath, "*.png", SearchOption.TopDirectoryOnly).OrderBy(filename => filename);
                foreach (string currentFile in Input_Filenames_PNG)
                {
                    Input_Filenames.Add(currentFile);
                }

                //  create a queue... 
                PopulateQueue(Input_Filenames);
                // start Threads... 
                StartThreads();
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        void StartThreads()
        {

            // TODO: Disable threading for initial version.. 

            processFileUsingZXing(0);
            return;

            /*

            Thread[] threads = new Thread[Number_Of_Threads];
            for (int i = 0; i < Number_Of_Threads; i++)
            {
                int LocalNum = i;
                WriteLog("Thread " + i + " created ....");
                threads[i] = new Thread(() => processFileUsingZXing(LocalNum)); // processFileUsingOpensource(LocalNum)); //processFileUsingDynamsoft(LocalNum)); // Program.processFile(LocalNum));
                threadcount++;
            }
            for (int i = 0; i < Number_Of_Threads; i++)
            {
                threads[i].Start();
                WriteLog("Thread " + i + " started ....");
            }
            for (int i = 0; i < Number_Of_Threads; i++)
            {
                threads[i].Join();
                WriteLog("Thread " + i + " joined ....");
            }

            */

        }

        /*
        * Not Used
        public  void MergeFiles(string pathFile1, string pathFile2, string pathResult)
        {
            File.WriteAllText(pathResult, File.ReadAllText(pathFile1) + File.ReadAllText(pathFile2));
        }
        */

        // include path in the fileName param... 
        public void CreateMetaDataFile(int threadID, string fileName, List<string> barCodeList, string parentFileName, string extractedData, string jsonContent)
        {


            try
            {
                string xmlFileExtension = ".daxml"; // document agent xml
                string jsonFileExtension = ".json"; // json file extension

                string xmlFileName = fileName + xmlFileExtension;
                string jsonFileName = fileName + jsonFileExtension;
                File.WriteAllText(jsonFileName, jsonContent);

                WriteLog("JSON file created: " + jsonFileName);

                // Check if the JSON file is created
                if (File.Exists(jsonFileName))
                {
                    WriteLog("JSON file confirmed to exist: " + jsonFileName);
                }
                else
                {
                    WriteLog("JSON file not found after creation attempt: " + jsonFileName);
                }


                // Create the XmlDocument.
                XmlDocument doc = new XmlDocument();
                doc.LoadXml("<document></document>");

                XmlElement newElem = doc.CreateElement("filename");
                newElem.InnerText = fileName.ToString();
                doc.DocumentElement.AppendChild(newElem);

                XmlAttribute newAttrib = doc.CreateAttribute("thread");
                newAttrib.InnerText = threadID.ToString();
                newElem.Attributes.Append(newAttrib);

                newAttrib = doc.CreateAttribute("original_filename");
                newAttrib.InnerText = parentFileName.ToString();
                newElem.Attributes.Append(newAttrib);

                XmlElement ocrElem = doc.CreateElement("OCRText");
                ocrElem.InnerText = extractedData;
                newElem.AppendChild(ocrElem);

                for (int i = 0; i < barCodeList.Count; i++)
                {
                    newAttrib = doc.CreateAttribute("barcode" + (i + 1).ToString());
                    newAttrib.InnerText = barCodeList[i];
                    newElem.Attributes.Append(newAttrib);
                }

                // Save the XML document to a file. White space is preserved (no white space).
                doc.PreserveWhitespace = true;
                doc.Save(xmlFileName);

               
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }



        //public void CreateMetaDataJsonFile(int threadID, string fileName, List<string> barCodeList, string parentFileName, string extractedData)
        //{
        //    try
        //    {
        //        string fileExtension = ".json"; // JSON file extension
        //        string newFileName = fileName + fileExtension;

        //        var metadata = new
        //        {
        //            FileName = fileName,
        //            ThreadID = threadID,
        //            OriginalFileName = parentFileName,
        //            OCRText = extractedData,
        //            Barcodes = barCodeList.Select((barcode, index) => new { Index = index + 1, Barcode = barcode }).ToDictionary(x => "barcode" + x.Index, x => x.Barcode)
        //        };

        //        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        //        string jsonString = JsonSerializer.Serialize(metadata, jsonOptions);

        //        File.WriteAllText(newFileName, jsonString);
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex.Message);
        //    }
        //}




        // checks if folder exists; if not, it creates it.. 
        public void CreateFolders()
        {
            try
            {
                Directory.CreateDirectory(Folder_Input);
                Directory.CreateDirectory(Folder_Input_Temp);
                Directory.CreateDirectory(Folder_Input_Processed);
                Directory.CreateDirectory(Folder_Output_Success);
                Directory.CreateDirectory(Folder_Input_Work);
                Directory.CreateDirectory(Folder_Output_Error);
                Directory.CreateDirectory(Folder_Scripts);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        // create the configuration file so the script can use... 
        public void CreateConfigFileForScript()
        {
            try
            {
                string NewFileName = "";
                NewFileName = Folder_Scripts + @"\config.daxml";

                // Create the XmlDocument.
                XmlDocument doc = new XmlDocument();
                doc.LoadXml("<configuration></configuration>");

                XmlElement newElem = doc.CreateElement("folderOutputError");
                newElem.InnerText = Folder_Output_Error.ToString();
                doc.DocumentElement.AppendChild(newElem);

                newElem = doc.CreateElement("folderOutputSuccess");
                newElem.InnerText = Folder_Output_Success.ToString();
                doc.DocumentElement.AppendChild(newElem);

                newElem = doc.CreateElement("folderInput");
                newElem.InnerText = Folder_Input.ToString();
                doc.DocumentElement.AppendChild(newElem);

                newElem = doc.CreateElement("folderInputWork");
                newElem.InnerText = Folder_Input_Work.ToString();
                doc.DocumentElement.AppendChild(newElem);

                newElem = doc.CreateElement("folderScripts");
                newElem.InnerText = Folder_Scripts.ToString();
                doc.DocumentElement.AppendChild(newElem);

                // Save the document to a file. White space is 
                // preserved (no white space).
                doc.PreserveWhitespace = true;
                doc.Save(NewFileName);

            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }

        }

        // Script will run in 32-bit mode. If Powershell script uses any database drivers, use the 32-bit database drivers. 
         public void RunScript()
        {
            // create Powershell runspace
            try
            {
                string[] FileList = Directory.GetFiles(Folder_Input_Work, "*.*");
                string FileName = Folder_Scripts + @"\processfiles.ps1";
                string ScriptText = System.IO.File.ReadAllText(FileName);
                // If no files to process, exit, no need to run script..
                if (FileList.Length >= 0)
                {
                    Runspace runspace = RunspaceFactory.CreateRunspace();

                    // open it
                    runspace.Open();

                    // create a pipeline and feed it the script text
                    Pipeline pipeline = runspace.CreatePipeline();
                    pipeline.Commands.AddScript(ScriptText);

                    pipeline.Commands.Add("Out-String");

                    // execute the script
                    Collection<PSObject> results = pipeline.Invoke();

                    // close the runspace
                    runspace.Close();

                    // convert the script result into a single string
                    StringBuilder stringBuilder = new StringBuilder();
                    foreach (PSObject obj in results)
                    {
                        stringBuilder.AppendLine(obj.ToString());
                    }
                }

            }
            catch (Exception ex)
            {
                string ErrorMessage = ex.Message;
                WriteLog(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "-" +  ErrorMessage);
            }
        }

        bool WriteLog(string textLine)
        {
            bool RetVal = false;

            if (LogFileName.Length == 0) LogFileName = Folder_Scripts + @"\service.log"; // by default if no log file passed in.. 

            System.IO.StreamWriter file = new System.IO.StreamWriter(LogFileName, true);
            file.WriteLine(DateTime.Now.ToLocalTime().ToString() + " " + textLine);
            file.Close();
            return (RetVal);
        }


        //public void TestFile(string filename, FILE_FORMAT fileFormat=FILE_FORMAT.TIFF)
        //{
        //    var barcodes_start_page = DecodeBarcodeZxing(filename, fileFormat);

        //}


        /**********************************************************
          PDFSHARP .NET IMAGE CODE                             
        **********************************************************/


        public void ExtractImageFromPDF(string filename)
        {

            PdfDocument document = PdfReader.Open(filename);

            int imageCount = 0;
            // Iterate pages
            foreach (PdfPage page in document.Pages)
            {
                // Get resources dictionary
                PdfDictionary resources = page.Elements.GetDictionary("/Resources");
                if (resources != null)
                {
                    // Get external objects dictionary
                    PdfDictionary xObjects = resources.Elements.GetDictionary("/XObject");
                    if (xObjects != null)
                    {

                        //PdfSharp.Pdf.PdfItem 
                        ICollection<PdfItem> items = xObjects.Elements.Values;
                        // Iterate references to external objects
                        foreach (PdfItem item in items)
                        {
                            PdfReference reference = item as PdfReference;
                            if (reference != null)
                            {
                                PdfDictionary xObject = reference.Value as PdfDictionary;
                                // Is external object an image?
                                if (xObject != null && xObject.Elements.GetString("/Subtype") == "/Image")
                                {
                                    ExportImage(xObject, ref imageCount);
                                }
                            }
                        }
                    }
                }
            }
            System.Windows.Forms.
            MessageBox.Show(imageCount + " images exported.", "Export Images");
        }

         void ExportImage(PdfDictionary image, ref int count)
        {
            string filter = image.Elements.GetName("/Filter");
            switch (filter)
            {
                case "/CCITTFaxDecode":
                    //PDFSharp.Extensions.

                case "/DCTDecode":
                    ExportJpegImage(image, ref count);
                    break;

                // Not Supported
                case "/FlateDecode":
                    ExportAsPngImage(image, ref count);
                    break;
            }
        }

         void ExportJpegImage(PdfDictionary image, ref int count)
        {
            // Fortunately JPEG has native support in PDF and exporting an image is just writing the stream to a file.
            byte[] stream = image.Stream.Value;
            FileStream fs = new FileStream(String.Format("Image{0}.jpeg", count++), FileMode.Create, FileAccess.Write);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(stream);
            bw.Close();
        }

         void ExportAsPngImage(PdfDictionary image, ref int count)
        {
            int width = image.Elements.GetInteger(PdfImage.Keys.Width);
            int height = image.Elements.GetInteger(PdfImage.Keys.Height);
            int bitsPerComponent = image.Elements.GetInteger(PdfImage.Keys.BitsPerComponent);

            // TODO: You can put the code here that converts vom PDF internal image format to a Windows bitmap
            // and use GDI+ to save it in PNG format.
            // It is the work of a day or two for the most important formats. Take a look at the file
            // PdfSharp.Pdf.Advanced/PdfImage.cs to see how we create the PDF image formats.
            // We don't need that feature at the moment and therefore will not implement it.
            // If you write the code for exporting images I would be pleased to publish it in a future release
            // of PDFsharp.
        }

        // check to see how many files have been processed for today.. 
        // Note: user can always just remove the daxml files to bypass the license check.. 
        int CheckHowManyFilesProcessedToday()
        {
            int FileCount = 0;
            try
            {
                var FileList = Directory.EnumerateFiles(Folder_Output_Success, "*.daxml", SearchOption.TopDirectoryOnly);
                foreach (string currentFile in FileList)
                {
                    DateTime FileDate = File.GetCreationTime(currentFile);
                    if (FileDate.Date == System.DateTime.Now.Date) FileCount++;
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            return (FileCount);
        }

        // Checks for license file. If none, found, assume it's a trial.. 
        // License check is based on index of where it can find certain keys. 
        // STANDARD, expects 9002 at column 5
        // PRO, expects 13529 at colum 7
        LICENSE_VERSION GetProductLicense(string licenseFile)
        {
            LICENSE_VERSION RetVal = LICENSE_VERSION.TRIAL;

            try
            {
                // check if file exists.
                if (File.Exists(licenseFile))
                {
                    // read to check if valid license file.. 
                    using (StreamReader sr = new StreamReader(licenseFile))
                    {
                        String line = sr.ReadToEnd();
                        if (line.IndexOf("9002") == 5)
                        {
                            RetVal = LICENSE_VERSION.STANDARD;
                        }
                        if (line.IndexOf("13529") == 5)
                        {
                            RetVal = LICENSE_VERSION.PRO;
                        }
                        if (line.IndexOf("723") == 5)
                        {
                            RetVal = LICENSE_VERSION.FREE;
                        }
                    }
                }
                else
                {
                    RetVal = LICENSE_VERSION.TRIAL;
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            return (RetVal);
        }

        // checks to see how many days have passed since the app was first used.. 
        // when app is first run, it will create a file with the trial date.. 
        // TODO: Test for UK or outside UK too. 
        int CheckTrialDays(string trialFile)
        {
            int NumberOfDaysPassed = 0;

            try
            {
                if (File.Exists(trialFile))
                {
                    // read to check if valid license file.. 
                    using (StreamReader sr = new StreamReader(trialFile))
                    {
                        String line = sr.ReadToEnd();
                        TimeSpan DateDelta = (System.DateTime.Now.Date- System.Convert.ToDateTime(line).Date);
                        NumberOfDaysPassed = DateDelta.Days;
                    }
                }
                else
                {
                    // create the file... 
                    NumberOfDaysPassed = 0;
                    System.IO.StreamWriter file = new System.IO.StreamWriter(trialFile, true);
                    file.WriteLine(System.DateTime.Now.Date.ToShortDateString());
                    file.Close();
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            return (NumberOfDaysPassed);
        }


    }
}
