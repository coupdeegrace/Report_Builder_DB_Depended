using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using reportBuilder.classes;
using System.Linq;

namespace reportBuilder
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct OpenFileName
    {
        public int lStructSize;
        public IntPtr hwndOwner;
        public IntPtr hInstance;
        public string lpstrFilter;
        public string lpstrCustomFilter;
        public int nMaxCustFilter;
        public int nFilterIndex;
        public string lpstrFile;
        public int nMaxFile;
        public string lpstrFileTitle;
        public int nMaxFileTitle;
        public string lpstrInitialDir;
        public string lpstrTitle;
        public int Flags;
        public short nFileOffset;
        public short nFileExtension;
        public string lpstrDefExt;
        public IntPtr lCustData;
        public IntPtr lpfnHook;
        public string lpTemplateName;
        public IntPtr pvReserved;
        public int dwReserved;
        public int flagsEx;
    }

    public class Program
    {
        [DllImport("comdlg32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool GetOpenFileName(ref OpenFileName ofn);

        private static string ShowDialog()
        {
            var ofn = new OpenFileName();
            ofn.lStructSize = Marshal.SizeOf(ofn);
            ofn.lpstrFilter = "CSV Files (*.csv)\0";
            ofn.lpstrFile = new string(new char[256]);
            ofn.nMaxFile = ofn.lpstrFile.Length;
            ofn.lpstrFileTitle = new string(new char[64]);
            ofn.nMaxFileTitle = ofn.lpstrFileTitle.Length;
            ofn.lpstrTitle = "Open File Dialog...";
            if (GetOpenFileName(ref ofn))
                return ofn.lpstrFile;
            return string.Empty;
        }

        public static void Main(string[] args)
        {
            
            var filename = ShowDialog();
            Console.WriteLine("Connecting to the database....");
            string connectionString = "Server=localhost;Database=ReportBuilderRU;User=sa;Password=Cyxariki_2404;Encrypt=True;TrustServerCertificate=True;";
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";",
                Encoding= Encoding.UTF8,
                HasHeaderRecord = true,
            };       
            try
            {   
                using (var reader = new StreamReader(filename))
                using (var csv = new CsvReader(reader, config))
                {
                    var records = csv.GetRecords<ReportClass>();
                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        Console.WriteLine("Complete!\nParsing your CSV into database....");
                        foreach (var record in records)
                        {
                            using (SqlCommand cmd = new SqlCommand("INSERT INTO ReportData (task_no, error_type, workspace, short_desc) VALUES (@task_no, @error_type, @workspace, @short_desc)", conn))
                            {
                                cmd.Parameters.AddWithValue("@task_no", record.task_no);
                                cmd.Parameters.AddWithValue("@error_type", record.error_type);
                                cmd.Parameters.AddWithValue("@workspace", record.workspace);
                                cmd.Parameters.AddWithValue(@"short_desc", record.short_desc);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }                
            } catch (Exception ex) { Console.WriteLine(ex); }
            Console.WriteLine("Complete!");
            compareData(connectionString);
            Console.ReadKey();
        }

        public static Dictionary<string, string> createErrorTypeDictionary(string dictionaryPath)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";",
                Encoding = Encoding.UTF8,
                HasHeaderRecord = true,
            };

            Dictionary<string, string> errorTypeDictionary = new Dictionary<string, string>();
            Console.WriteLine("Reading your dictionary....");

            using (var reader = new StreamReader(dictionaryPath, Encoding.UTF8))
            using (var csv = new CsvReader(reader, config))
            {
                var records = csv.GetRecords<errorTypeDictionaryClass>();
                try
                {
                    foreach (var record in records)
                    {
                        errorTypeDictionary[record.errorType] = record.specification;
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex);}             
            }
            Console.WriteLine("Complete!");
            return errorTypeDictionary;            
        }

        public static void compareData(string connectionString)
        {
            string dictionaryPath = @"C:\Users\avglushkov\Desktop\автоматизация\dictionary.csv";
            Dictionary<string, string> categorizedReportData =  createErrorTypeDictionary(dictionaryPath);

            string query = @"select error_type, short_desc from ReportData";

            Dictionary<string, int> errorTypeCounts = new Dictionary<string, int>();

            using (SqlConnection connection = new SqlConnection(connectionString)) 
            {
                connection.Open();

                using (SqlCommand cmd = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader()) 
                    {
                        while (reader.Read())
                        {
                            string errorType = reader["error_type"] as string;
                            string shortDesc = reader["short_desc"] as string;

                            string comparisonKey = !string.IsNullOrEmpty(errorType) ? errorType : shortDesc;

                            if (categorizedReportData.TryGetValue(comparisonKey, out string desc))
                            {
                                if (errorTypeCounts.ContainsKey(desc))
                                {
                                    errorTypeCounts[desc]++;
                                }
                                else
                                {
                                    errorTypeCounts[desc] = 1;
                                }
                            }
                        }
                    }
                }
            }
            foreach (var kvp in errorTypeCounts)
            {
                Console.WriteLine($"Description: {kvp.Key}, Count: {kvp.Value}");
            }
        }
    }
}