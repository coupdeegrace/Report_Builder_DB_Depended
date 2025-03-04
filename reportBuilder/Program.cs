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
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.Extensions.Configuration;

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

        public static string GetConnectionString()
        {
            var configuration = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

            return configuration.GetConnectionString("DefaultConnection");
        }

        public static void ClearUpTable(string connectionString, string tableName)
        {
            string query = $"delete from {tableName} where task_no is not null";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                   int affectedRows = cmd.ExecuteNonQuery();
                   Console.WriteLine($"\nCleared up {affectedRows} rows");
                }
            }
        }

        public static void Main(string[] args)
        {
            var filename = ShowDialog();
            Console.WriteLine("Connecting to the database....");
            string connectionString = GetConnectionString();
            Console.Write("Complete!");

            string tableName = "ReportData";
            Console.Write("Clearing up your database....");
            ClearUpTable(connectionString, tableName);

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
                        Match match = Regex.Match(record.errorType, @"[A-Z]+-\d+");
                        if (match.Success)                        
                            errorTypeDictionary[record.errorType] = record.specification;
                        else 
                            errorTypeDictionary[record.errorType.ToLower()] = record.specification;
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex);}             
            }
            Console.WriteLine("Complete!");
            return errorTypeDictionary;            
        }

        public static string NormalizeErrorType(string errorType)
        {
            Match match = Regex.Match(errorType, @"[A-Z]+-\d+");

            if (match.Success)
            {
                return match.Value;
            }
            string result = Regex.Replace(errorType, @"\s *\(.*?\)\s*", "");
            return result.Trim().ToLower();
        }

        public static string NoramlizeShortDesc(string shortDesc)
        {
            string result = Regex.Replace(shortDesc, @"\s *\(.*?\)\s*", "");
            return result.Trim().ToLower();
        }

        public static void compareData(string connectionString)
        {
            Console.WriteLine("Please, select your dictionary....");
            string dictionaryPath = ShowDialog();
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
                        int taskAmount = 0;
                        List<string> missingErrorTypes = new List<string>();

                        while (reader.Read())
                        {                            
                            string errorType = reader["error_type"] as string;
                            string shortDesc = reader["short_desc"] as string;

                            string normalizedErrorType = NormalizeErrorType(errorType);
                            string normalizedDesc = NoramlizeShortDesc(shortDesc);
                            string comparisonKey = !string.IsNullOrEmpty(normalizedErrorType) ? normalizedErrorType : normalizedDesc;

                            if (categorizedReportData.TryGetValue(comparisonKey, out string desc))
                            {
                                if (errorTypeCounts.ContainsKey(desc))
                                {
                                    errorTypeCounts[desc]++;
                                    taskAmount++;
                                }
                                else
                                {
                                    errorTypeCounts[desc] = 1;
                                    taskAmount++;
                                }
                            }
                            else
                                missingErrorTypes.Add(comparisonKey);
                        }

                        Console.WriteLine($"Amount of tasks: {taskAmount}");

                        if (missingErrorTypes.Count > 0)
                        {
                            foreach (var error in missingErrorTypes)
                            {
                                Console.WriteLine($"{error} was not found in dictionary");
                            }
                        }
                        else
                            Console.WriteLine("All error types was found in dictionary. No updates needed");
                    }
                }
            }
            string credentialsPath = "C:\\Users\\avglushkov\\Downloads\\credentials.json";
            string spreadsheetId = "1Y_wPVvbCL79Drtc9dyIFctPCFPhTnLhEvRS5Ue3pCfM";

            GoogleSheetsService googleSheetsService = new GoogleSheetsService(credentialsPath, spreadsheetId);

            googleSheetsService.InsertData(errorTypeCounts);
        }
    }
}