using CsvLib;
using ExifLibrary;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace PhotoGPS
{
    class Program
    {
        //contain sub folder
        public static List<FileInfo> GetFiles(DirectoryInfo directoryInfo)
        {
            var fileList = new List<FileInfo>();
            fileList.AddRange(directoryInfo.GetFiles());//searchPattern:"*.jpg"
            foreach (var childDirectory in directoryInfo.GetDirectories())
            {
                fileList.AddRange(GetFiles(childDirectory));
            }
            return fileList;
        }

        public static void SaveToCSV(DataTable dt, string fileName = "GeoInfos.csv")
        {
            CsvStreamWriter csv = new CsvStreamWriter();
            csv.AddData(dt);
            csv.Save(System.Environment.CurrentDirectory + @"\" + fileName, Encoding.UTF8);
        }

        public static DataTable GetGeoInfos()
        {
            DataTable dt = new DataTable();

            //prepare data struct for DataTable
            dt.Columns.Add("fileName", System.Type.GetType("System.String"));
            dt.Columns.Add("Latitude", System.Type.GetType("System.String"));
            dt.Columns.Add("Longitude", System.Type.GetType("System.String"));
            dt.Columns.Add("GPSLatitudeRef", System.Type.GetType("System.String"));
            dt.Columns.Add("GPSLatitude", System.Type.GetType("System.String"));
            dt.Columns.Add("GPSLongitudeRef", System.Type.GetType("System.String"));
            dt.Columns.Add("GPSLongitude", System.Type.GetType("System.String"));
            dt.Columns.Add("GPSAltitudeRef", System.Type.GetType("System.String"));
            dt.Columns.Add("GPSAltitude", System.Type.GetType("System.String"));

            // add titles on first line
            DataRow firstRow = dt.NewRow();
            firstRow["fileName"] = "fileName";
            firstRow["Latitude"] = "Latitude";
            firstRow["Longitude"] = "Longitude";
            firstRow["GPSLatitudeRef"] = "GPSLatitudeRef";
            firstRow["GPSLatitude"] = "GPSLatitude";
            firstRow["GPSLongitudeRef"] = "GPSLongitudeRef";
            firstRow["GPSLongitude"] = "GPSLongitude";
            firstRow["GPSAltitudeRef"] = "GPSAltitudeRef";
            firstRow["GPSAltitude"] = "GPSAltitude";
            dt.Rows.Add(firstRow);

            DirectoryInfo directoryInfo = new DirectoryInfo(System.Environment.CurrentDirectory);
            var fileInfos = GetFiles(directoryInfo);

            int fileCount = fileInfos.Count;
            int succeedCount = 0;
            int noInfoCount = 0;
            int exceptionCount = 0;
            foreach (var fileInfo in fileInfos)
            {
                Console.WriteLine(fileInfo);
                ImageFile file = null;
                try
                {
                    file = ImageFile.FromFile(fileInfo.FullName);
                    GPSLatitudeLongitude latitude = file.Properties[ExifTag.GPSLatitude] as GPSLatitudeLongitude;
                    GPSLatitudeLongitude longitude = file.Properties[ExifTag.GPSLongitude] as GPSLatitudeLongitude;

                    DataRow dr = dt.NewRow();
                    dr["fileName"] = fileInfo.Name;
                    dr["Latitude"] = latitude.ToFloat();
                    dr["Longitude"] = longitude.ToFloat();
                    dr["GPSLatitudeRef"] = file.Properties[ExifTag.GPSLatitudeRef];
                    dr["GPSLatitude"] = file.Properties[ExifTag.GPSLatitude];
                    dr["GPSLongitudeRef"] = file.Properties[ExifTag.GPSLongitudeRef];
                    dr["GPSLongitude"] = file.Properties[ExifTag.GPSLongitude];
                    dr["GPSAltitudeRef"] = file.Properties[ExifTag.GPSAltitudeRef];
                    dr["GPSAltitude"] = file.Properties[ExifTag.GPSAltitude];
                    dt.Rows.Add(dr);

                    succeedCount++;
                }
                catch (KeyNotFoundException e)
                {
                    Console.WriteLine("Info Not Found  file:" + fileInfo.FullName + " exception:" + e);
                    noInfoCount++;
                }
                catch (Exception e)
                {
                    Console.WriteLine("Other Exception file:" + fileInfo.FullName + " exception:" + e);
                    exceptionCount++;
                }
            }
            Console.WriteLine("fileCount:\t" + fileCount);
            Console.WriteLine("succeedCount:\t" + succeedCount);
            Console.WriteLine("noInfoCount:\t" + noInfoCount);
            Console.WriteLine("exceptionCount:\t" + exceptionCount);
            return dt;
        }

        [STAThread()]
        static void Main(string[] args)
        {
            //setConsoleWindowVisibility(false, Console.Title);

            var dt = GetGeoInfos();
            SaveToCSV(dt, args.Length == 0 ? "GeoInfos.csv" : args[0]);
        }


        #region tool for hide console
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        public static void setConsoleWindowVisibility(bool visible, string title)
        {
            // below is Brandon's code            
            //Sometimes System.Windows.Forms.Application.ExecutablePath works for the caption depending on the system you are running under.           
            IntPtr hWnd = FindWindow(null, title);

            if (hWnd != IntPtr.Zero)
            {
                if (!visible)
                    //Hide the window                    
                    ShowWindow(hWnd, 0); // 0 = SW_HIDE                
                else
                    //Show window again                    
                    ShowWindow(hWnd, 1); //1 = SW_SHOWNORMA           
            }
        }
        #endregion
    }
}
