
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.IO;

internal class Program
{

    private static void Main(string[] args)
    {
        //SiteData();
        //LoadAddress();
        //GetCoord();
        //SeedBTS();
        //SeedClient();
        //SeedCircuit();
        SeedMPLSPoP();        
    }

    private static void SeedMPLSPoP()
    {
        string XlFileName = @"C:\ENTERPRISEDATA\Book101.xlsx";

        MPLSPoP mplspop = new MPLSPoP();

        string s = string.Empty;

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage data = new ExcelPackage(new FileInfo(XlFileName)))
            {
                ExcelWorksheet worksheet = data.Workbook.Worksheets[0];//The first sheet on the workbook
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                for (int row = 1; row <= end.Row; row++)//Row 1 contains titles   //worksheet.Dimension.End.Row
                {
                    for (int col = start.Column; col <= 1; col++)
                    {
                        mplspop.Id = (worksheet.Cells[row, 1].Value).ToString();
                        mplspop.BTS = (worksheet.Cells[row, 2].Value).ToString();
                        mplspop.NEName = (worksheet.Cells[row, 3].Value).ToString();
                        mplspop.NEType = (worksheet.Cells[row, 4].Value).ToString();
                        mplspop.NEIpAddress = (worksheet.Cells[row, 5].Value).ToString();
                        

                        s += $"new MPLSPoP{{Id={mplspop.Id}, BTSId={mplspop.BTS}, NEName=\"{mplspop.NEName}\", NEType=\"{mplspop.NEType}\", NEIpAddress=\"{mplspop.NEIpAddress}\"}},\n";
                    }
                }
                Console.WriteLine(s);
                Console.ReadKey();
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error" + ex.ToString());
        }
    }


    private static void SeedCircuit()
    {
        string XlFileName = @"C:\ENTERPRISEDATA\Book101.xlsx";

        Circuit circuit = new Circuit();
        string s = string.Empty;

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage data = new ExcelPackage(new FileInfo(XlFileName)))
            {
                ExcelWorksheet worksheet = data.Workbook.Worksheets[5];//The first sheet on the workbook
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                for (int row = 1; row <= end.Row; row++)//Row 1 contains titles   //worksheet.Dimension.End.Row
                {
                    for (int col = start.Column; col <= 1; col++)
                    {
                        circuit.Id = (worksheet.Cells[row, 1].Value).ToString();
                        circuit.CircuitRef = (worksheet.Cells[row, 2].Value).ToString();
                        circuit.ClientId = (worksheet.Cells[row, 3].Value).ToString();
                        circuit.CircuitName = (worksheet.Cells[row, 4].Value).ToString();
                        circuit.Address = (worksheet.Cells[row, 5].Value).ToString();
                        circuit.Town = (worksheet.Cells[row, 6].Value).ToString();
                        circuit.StateId = (worksheet.Cells[row, 7].Value).ToString();
                        circuit.Latitude = (worksheet.Cells[row, 8].Value).ToString();
                        circuit.Longitude = (worksheet.Cells[row, 9].Value).ToString();
                        circuit.Coordinates = (worksheet.Cells[row, 10].Value).ToString();
                        circuit.ServiceType = (worksheet.Cells[row, 11].Value).ToString();
                        circuit.JCCApprovedDate = (worksheet.Cells[row, 12].Value).ToString();
                        circuit.AnnualRevenue = (worksheet.Cells[row, 13].Value).ToString();
                        circuit.Bandwidth = (worksheet.Cells[row, 14].Value).ToString();
                        circuit.CircuitState = (worksheet.Cells[row, 15].Value).ToString();
                        circuit.AccountManager = (worksheet.Cells[row, 16].Value).ToString();
                        circuit.ProjectManager = (worksheet.Cells[row, 17].Value).ToString();
                        circuit.TAM = (worksheet.Cells[row, 18].Value).ToString();
                        s += $"new Circuit{{Id={circuit.Id}, CircuitRef={circuit.CircuitRef}, ClientId={circuit.ClientId}, CircuitName=\"{circuit.CircuitName}\", Address=\"{circuit.Address}\", Town=\"{circuit.Town}\", StateId={circuit.StateId}, Latitude={circuit.Latitude}, Longitude={circuit.Longitude}, Coordinates=\"{circuit.Coordinates}\", ServiceType={circuit.ServiceType}, JCCApprovedDate=DateOnly.FromDateTime(DateTime.Parse(\"{circuit.JCCApprovedDate}\")), AnnualRevenue={circuit.AnnualRevenue}, Bandwidth={circuit.Bandwidth}, CircuitState={circuit.CircuitState}, AccountManager=\"{circuit.AccountManager}\", ProjectManager=\"{circuit.ProjectManager}\", TAM=\"{circuit.TAM}\"}},\n";
                    }
                }
                Console.WriteLine(s);
                Console.ReadKey();
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error" + ex.ToString());
        }
    }

    private static void SeedClient()
    {
        string XlFileName = @"C:\Users\Seyi\Desktop\ENTERPRISEDATA\Read from xl\Book101.xlsx";

        ClientData clientData = new ClientData();
        string s = string.Empty;

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage data = new ExcelPackage(new FileInfo(XlFileName)))
            {
                ExcelWorksheet worksheet = data.Workbook.Worksheets[1];//The first sheet on the workbook
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                for (int row = 1; row <= end.Row; row++)//Row 1 contains titles   //worksheet.Dimension.End.Row
                {
                    for (int col = start.Column; col <= 1; col++)
                    {
                        clientData.Id = (worksheet.Cells[row, 1].Value).ToString();
                        clientData.Ref = (worksheet.Cells[row, 2].Value).ToString();
                        clientData.Name = (worksheet.Cells[row, 3].Value).ToString();
                        s += $"new Client{{Id={clientData.Id}, ClientRef={clientData.Ref}, ClientName=\"{clientData.Name}\"}},\n";
                    }
                }
                Console.WriteLine(s);
                Console.ReadKey();
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error" + ex.ToString());
        }
    }

    private static void SeedBTS()
    {

        string XlFileName = @"C:\Users\Seyi\Desktop\ENTERPRISE DATA\Read from xl\Book101.xlsx";

        ReadToSeed readToSeed = new ReadToSeed();
        string s = string.Empty;

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage data = new ExcelPackage(new FileInfo(XlFileName)))
            {
                ExcelWorksheet worksheet = data.Workbook.Worksheets[1];//The first sheet on the workbook
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                for (int row = 1; row <= end.Row; row++)//Row 1 contains titles   //worksheet.Dimension.End.Row
                {
                    for (int col = start.Column; col <= 1; col++)
                    {
                        readToSeed.Id = (worksheet.Cells[row, 1].Value).ToString();
                        readToSeed.BTSName = (worksheet.Cells[row, 2].Value).ToString();
                        readToSeed.LocationAddress = (worksheet.Cells[row, 3].Value).ToString();
                        readToSeed.StateId = (worksheet.Cells[row, 4].Value).ToString();
                        readToSeed.Latitude = (worksheet.Cells[row, 5].Value).ToString();
                        readToSeed.Longitude = (worksheet.Cells[row, 6].Value).ToString();
                        readToSeed.Coordinates = (worksheet.Cells[row, 7].Value).ToString();
                        s += $"new BTS{{Id={readToSeed.Id}, BTSName=\"{readToSeed.BTSName}\", LocationAddress=\"{readToSeed.LocationAddress}\", StateId={readToSeed.StateId}, Latitude={readToSeed.Latitude}, Longitude={readToSeed.Longitude}, Coordinates=\"{readToSeed.Coordinates}\"}},\n";
                    }
                }
                Console.WriteLine(s);
                Console.ReadKey();
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error" + ex.ToString());
        }
        //return SiteInfo;
    }

    private static void GetCoord()
    {
        
        string XlFileName = @"C:\Users\Seyi\Desktop\ENTERPRISE DATA\Read from xl\Book101.xlsx";
        
        Coord coord = new Coord();
        string s = string.Empty;

        try 
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (ExcelPackage data = new ExcelPackage(new FileInfo(XlFileName)))
            {
                ExcelWorksheet worksheet = data.Workbook.Worksheets[0];//The first sheet on the workbook
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                for (int row = 1; row <= end.Row; row++)//Row 1 contains titles   //worksheet.Dimension.End.Row
                {
                    for (int col = start.Column; col <= 1; col++)
                    {
                        coord.Latitude = Convert.ToDouble(worksheet.Cells[row, 1].Value);
                        coord.Longitude = Convert.ToDouble(worksheet.Cells[row, 2].Value);
                        s += coord.Latitude.ToString()+"-"+coord.Longitude.ToString()+"+"+coord.Coordinates + "\n";
                    }
                }
                Console.WriteLine(s);
                Console.ReadKey();
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error" + ex.ToString());
        }
        //return SiteInfo;
    }

    private static void SiteData()
    {
        //string XlFileName = @"C:\Users\Seyi\Desktop\Glentek\Site Reconciliation.xlsx";

        string XlFileName = @"C:\Users\Seyi\Desktop\Glentek\EnterpriseData.xlsx";
        string note = @"C:\Users\Seyi\Desktop\Glentek\all.txt";
        Site site = new Site();
        StreamWriter s = new StreamWriter(note,true);

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            

            #region Dialog to Get Filename
            //sheet = new OpenFileDialog()
            //{
            //    Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls",
            //    Title = "Select Excel File"
            //};

            //if (sheet.ShowDialog() == DialogResult.OK)
            //{
            //    XlFileName = sheet.FileName;
            //}
            #endregion
            

            using (ExcelPackage data = new ExcelPackage(new FileInfo(XlFileName)))
            {
                ExcelWorksheet worksheet = data.Workbook.Worksheets[4];//The first sheet on the workbook
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                //int startRow;
                //int endRow;

                for (int row = 1; row <= end.Row; row++)//Row 1 contains titles   //worksheet.Dimension.End.Row
                {
                    
                    for (int col = start.Column; col <= 1; col++)
                    {
                        site.BTS = worksheet.Cells[row, 1].Value.ToString();
                        site.Address = worksheet.Cells[row, 2].Value.ToString();
                        //site.State = worksheet.Cells[row, 3].Value.ToString();
                        //site.Latitude = worksheet.Cells[row, 4].Value.ToString();
                        //site.Longitude = worksheet.Cells[row, 5].Value.ToString();
                        //customer.lastmileContractor = worksheet.Cells[row, 12].Value != null ? worksheet.Cells[row, 12].Value.ToString() : "no value";
                    }
                    //SiteInfo.Add(site);

                    //s.WriteLine($"insert into dbo.BTS values('{site.BTS}','{site.Address}',{site.State},'{site.Latitude}','{site.Longitude}');");
                    s.WriteLine($"insert into dbo.Customer values('{site.BTS}',{site.Address});");
                }
            }

            #region Separate the good from the bad rows
            #endregion

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error" + ex.ToString());
        }
        //return SiteInfo;
    }

    private static void LoadAddress()
    {
        string note = @"C:\Users\Seyi\Desktop\Glentek\all.txt";
        StreamWriter s = new StreamWriter(note, true);
        string XlFileName = @"C:\Users\Seyi\Desktop\Glentek\EnterpriseData.xlsx";
        SomeInfo someInfo = new SomeInfo();
        AddressInfo addressInfo = new AddressInfo();
        List<SomeInfo> some; 
        List<AddressInfo> address;

        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage data = new ExcelPackage(new FileInfo(XlFileName)))
            {
                ExcelWorksheet worksheet = data.Workbook.Worksheets[0];//The first sheet on the workbook
                ExcelWorksheet workshitu = data.Workbook.Worksheets[1];//The second sheet on the workbook
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                var startu = workshitu.Dimension.Start;
                var entu = workshitu.Dimension.End;
                some = new List<SomeInfo>();
                address = new List<AddressInfo>();


                for (int row = 2; row <= end.Row; row++)//Row 1 contains titles   //worksheet.Dimension.End.Row
                {
                    for (int col = start.Column; col <= 1; col++)
                    {
                        someInfo.siteName = worksheet.Cells[row, 1].Value.ToString();
                        //someInfo.addy = worksheet.Cells[row, 2].Value.ToString();
                        
                    }
                }

                for (int row = 2; row <= entu.Row; row++)//Row 1 contains titles   //worksheet.Dimension.End.Row
                {
                    for (int col = start.Column; col <= 2; col++)
                    {
                        addressInfo.siteName = workshitu.Cells[row, 1].Value.ToString();
                        addressInfo.addy = workshitu.Cells[row, 2].Value.ToString();                        
                    }
                    address.Add(addressInfo);
                }
            }

            #region Separate the good from the bad rows
            #endregion



        }
        catch (Exception ex)
        {
            Console.WriteLine("Error" + ex.ToString());
        }
    }

}



public class MPLSPoP
{
    public string Id { get; set; }

    public string BTS { get; set; }

    public string NEName { get; set; }

    public string NEType { get; set; }
    public string NEIpAddress { get; set; }
}


public class SomeInfo
{
    public string? siteName { get; set; }
    public string? addy { get; set; }
}

public class AddressInfo
{
    public string? siteName { get; set; }
    public string? addy { get; set; }
}

public class Coord
{
    public double? Latitude { get; set; }
    public double? Longitude { get; set; }
    public string? Coordinates
    {
        get => CalculateCoord();
    }

    private string? CalculateCoord()
    {

        string latitudeCoord, longitudeCoord;
        if (Latitude.HasValue && Longitude.HasValue)
        {
            int degrees;
            double minutes, seconds;
            // set decimal_degrees value here
            if (Latitude.Value.ToString().IndexOf('.') != -1)
            {
                degrees = Convert.ToInt32(Latitude.Value.ToString().Split(".")[0]);
                minutes = (Latitude.Value - degrees) * 60;
                seconds = (minutes - Math.Floor(minutes)) * 60.0;

                // get rid of fractional part
                minutes = Math.Floor(minutes);
                seconds = Math.Round(seconds, 2);
                latitudeCoord = $"{degrees}\u00b0{minutes}'{seconds}\"N";
            }
            else
            {
                latitudeCoord = $"{Latitude.Value}\u00b00'0\"N";
            }

            if (Longitude.Value.ToString().IndexOf('.') != -1)
            {
                degrees = Convert.ToInt32(Longitude.Value.ToString().Split(".")[0]);
                minutes = (Longitude.Value - degrees) * 60;
                seconds = (minutes - Math.Floor(minutes)) * 60.0;

                // get rid of fractional part
                minutes = Math.Floor(minutes);
                seconds = Math.Round(seconds, 2);
                longitudeCoord = $"{degrees}\u00b0{minutes}'{seconds}\"E";
            }
            else
            {
                longitudeCoord = $"{Latitude.Value}\u00b00'0\"E";
            }
            return $"{latitudeCoord} {longitudeCoord}";
        }
        else
        {
            return null;
        }
    }
}

public class Site
{
    public string? BTS { get; set; }
    public string? Address { get; set; }
    public string? State { get; set; }
    public string? Latitude { get; set; }
    public string? Longitude { get; set; }


//    public int num(string bsc)
//    {

//        return bsc == "Abia" ? 1 : bsc == "Adamawa" ? 2 :
//bsc == "Akwa-Ibom" ? 3 :
//bsc == "Anambra" ? 4 :
//bsc == "Bauchi" ? 5 :
//bsc == "Bayelsa" ? 6 :
//bsc == "Benue" ? 7 :
//bsc == "Borno" ? 8 :
//bsc == "Cross River" ? 9 :
//bsc == "Delta" ? 10 :
//bsc == "Ebonyi" ? 11 :
//bsc == "Edo" ? 12 :
//bsc == "Ekiti" ? 13 :
//bsc == "Enugu" ? 14 :
//bsc == "Gombe" ? 15 :
//bsc == "Imo" ? 16 :
//bsc == "Jigawa" ? 17 :
//bsc == "Kaduna" ? 18 :
//bsc == "Kano" ? 19 :
//bsc == "Katsina" ? 20 :
//bsc == "Kebbi" ? 21 :
//bsc == "Kogi" ? 22 :
//bsc == "Kwara" ? 23 :
//bsc == "Lagos" ? 24 :
//bsc == "Nasarawa" ? 25 :
//bsc == "Niger" ? 26 :
//bsc == "Ogun" ? 27 :
//bsc == "Ondo" ? 28 :
//bsc == "Osun" ? 29 :
//bsc == "Oyo" ? 30 :
//bsc == "Plateau" ? 31 :
//bsc == "Rivers" ? 32 :
//bsc == "Sokoto" ? 33 :
//bsc == "Taraba" ? 34 :
//bsc == "Yobe" ? 35 :
//bsc == "Zamfara" ? 36 :
//bsc == "Abuja" ? 37 : 0;

//    }

}

public class ReadToSeed
{
    public string? Id { get; set; }

    public string? BTSName { get; set; }

    public string? LocationAddress { get; set; }

    public string? StateId { get; set; }

    public string? Latitude { get; set; }

    public string? Longitude { get; set; }
   
    public string? Coordinates { get; set; }

    
}

public class ClientData
{
    public string Id { get; set; }
    public string Ref { get; set; }
    public string Name { get; set; }
}

public class Circuit
{
    public string Id { get; set; }
    
    public string CircuitRef { get; set; }

    public string ClientId { get; set; }

    public string CircuitName { get; set; }

    public string Address { get; set; }

    public string Town { get; set; }

    public string StateId { get; set; }
    
    public string Latitude { get; set; }

    public string Longitude { get; set; }

    public string? Coordinates { get; set; }

    public string ServiceType { get; set; }

    public string JCCApprovedDate { get; set; }

    public string AnnualRevenue { get; set; }

    public string Bandwidth { get; set; }

    public string CircuitState { get; set; }

    public string AccountManager { get; set; }
    public string ProjectManager { get; set; }
    public string TAM { get; set; }
}