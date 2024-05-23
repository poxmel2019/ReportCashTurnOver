using Microsoft.Data.SqlClient;
using Spire.Xls;
using System.Security.Cryptography.Pkcs;

namespace ReportCashTurnOver
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Workbook workbook = new Workbook();

            Worksheet worksheet = workbook.Worksheets[0];

            string connectionString = "Server=hbmssqltest.halykbank.nb;Database=CorePayments;User ID=CorePayments;Password=0coayiwbYVReR;TrustServerCertificate=true;";

            // year
            Console.WriteLine("Year:");
            int year = 0;
            bool isCorrectFormat = false;
            while (!isCorrectFormat)
            {
                try
                {
                    year = Convert.ToInt32(Console.ReadLine());
                    isCorrectFormat = true;

                }
                catch (FormatException ex)
                {
                    Console.WriteLine("Incorrect format!");
                }
            }

            //month
            Dictionary<int, string> months = new Dictionary<int, string>()
            {
                {1, "january"},
                {2, "february"},
                {3, "march"},
                {4, "april"},
                {5, "may"},
                {6, "june"},
                {7, "july"},
                {8, "august"},
                {9, "september"},
                {10, "october"},
                {11, "november"},
                {12, "december"}
            };

            Console.WriteLine("Month:");
            
            int month = 0;
            isCorrectFormat = false;

            while (!isCorrectFormat)
            {
                try
                {
                    month = Convert.ToInt32(Console.ReadLine());
                    if (month > 0 && month < 13) 
                    {
                        Console.WriteLine($"{months[month]}");
                        isCorrectFormat = true;
                    }
                    else
                    {
                        Console.WriteLine("Out of range months!");
                    }
                    

                }
                catch (FormatException ex)
                {
                    Console.WriteLine("Incorrect format!");
                }
            }

            string sqlExpression = "select \r\n" +                          //81
                "s.Name as ServiceName," +
                "s.Description as 'Наименование сервиса', " +
                "\r\ncount(p.Id) as 'Количество платежей', \r\n" +
                "sum(cast(pp.Value as numeric)) as 'Сумма платежей'\r\n" +
                "from Processes (nolock) p\r\n" +
                "join Services (nolock) s \r\n" +
                "on s.Id = p.ServiceId\r\n" +
                "join \r\n" +
                "(\r\nselect ProcessId, pp.Value \r\n" +
                "from Processes (nolock) p\r\n" +
                "join ProcessProperties (nolock) pp " +
                "on pp.ProcessId = p.Id\r\n" +
                "where StartDate " +
                $"between '{year}-{month}-01' " +
                $"and '{year}-{(month+1)}-01'\r\n" +
                "and LastPortalServiceOperationStateId in (16,37)\r\n" +
                "and pp.ServicePropertyName = 'amount'\r\n) pp " +
                "on pp.ProcessId = p.Id\r\n" +
                "group by s.Name, s.Description\r\n" +
                "order by s.Name \r\n;\r\n";
            Console.WriteLine(sqlExpression);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                await connection.OpenAsync();

                SqlCommand command = new SqlCommand(sqlExpression, connection);
                SqlDataReader reader = await command.ExecuteReaderAsync();

                if (reader.HasRows)
                {
                    string columnName1 = reader.GetName(0);
                    string columnName2 = reader.GetName(1);
                    string columnName3 = reader.GetName(2);
                    string columnName4 = reader.GetName(3);


                    worksheet.Range[1, 1].Value = columnName1;
                    worksheet.Range[1, 2].Value = columnName2;
                    worksheet.Range[1, 3].Value = columnName3;
                    worksheet.Range[1, 4].Value = columnName4;

                    Console.WriteLine($"{columnName1}\t{columnName2}");

                    int i = 2;

                    while (await reader.ReadAsync())
                    {
                        object serviceName = reader.GetValue(0);
                        object nameService = reader.GetValue(1);
                        object payQuantity = reader.GetValue(2);
                        object paySum = reader.GetValue(3);

                        Console.WriteLine($"{serviceName}\t{nameService}\t{payQuantity}\t{paySum}");
                        string[] array = { serviceName.ToString(), nameService.ToString(), payQuantity.ToString(), paySum.ToString()};
                        worksheet.InsertArray(array, i, 1, false);

                        i++;
                    }

                    // using style to first string
                    CellStyle style = workbook.Styles.Add("newStyle");
                    style.Font.IsBold = true;
                    worksheet.Range[1, 1, 1, 6].Style = style;

                    // fit width of columns
                    worksheet.AllocatedRange.AutoFitColumns();

                    string file_name = $"ReportCashTurnOver_{months[month]}_{year.ToString()}";

                    // save to excel file
                    try
                    {
                        workbook.SaveToFile($"C:\\for_work\\code\\my_projects\\ForExcel\\{file_name}.xlsx");

                    }
                    catch (IOException)
                    {
                        Console.WriteLine("The file is busy!");
                    }

                }

                await reader.CloseAsync();

            }

            Console.Read();
        }


    }

}


