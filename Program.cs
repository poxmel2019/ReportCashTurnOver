using Microsoft.Data.SqlClient;
using Spire.Xls;

namespace ReportCashTurnOver
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Workbook workbook = new Workbook();

            Worksheet worksheet = workbook.Worksheets[0];

            string connectionString = "Server=hbmssqltest.halykbank.nb;Database=CorePayments;User ID=CorePayments;Password=0coayiwbYVReR;TrustServerCertificate=true;";
            Console.WriteLine(connectionString);

            string sqlExpression = "select ProcessId, pp.Value \r\n" +
                "from Processes p\r\n" +
                "join ProcessProperties pp " +
                "on pp.ProcessId = p.Id\r\n" +
                "where StartDate between '2024-04-01' and '2024-05-01'\r\n" +
                "and LastPortalServiceOperationStateId " +
                "in (16,37)\r\n" +
                "and pp.ServicePropertyName = 'amount'";
            //string sqlExpression = "select id, name from portals order by id";
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

                    worksheet.Range[1, 1].Value = columnName1;
                    worksheet.Range[1, 2].Value = columnName2;

                    Console.WriteLine($"{columnName1}\t{columnName2}");

                    int i = 2;

                    while (await reader.ReadAsync())
                    {
                        object processId = reader.GetValue(0);
                        object value = reader.GetValue(1);
                        Console.WriteLine($"{processId}\t{value}");
                        string[] array = { processId.ToString(), value.ToString() };
                        worksheet.InsertArray(array, i, 1, false);

                        i++;
                    }

                    // using style to first string
                    CellStyle style = workbook.Styles.Add("newStyle");
                    style.Font.IsBold = true;
                    worksheet.Range[1, 1, 1, 6].Style = style;

                    // fit width of columns
                    worksheet.AllocatedRange.AutoFitColumns();

                    // save to excel file
                    workbook.SaveToFile("C:\\for_work\\code\\my_projects\\ForExcel\\ReportCashTurnOver.xlsx");

                }

                await reader.CloseAsync();

            }

            Console.Read();
        }


    }

}


