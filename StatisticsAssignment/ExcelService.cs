using Microsoft.EntityFrameworkCore;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using StatisticsAssignment.Db;
using StatisticsAssignment.Db.Entities;

namespace StatisticsAssignment
{
    public sealed class ExcelService
    {
        private readonly AssignmentDbContext dbContext;

        private const int CountryColumnIndex = 1;
        private const int _2002_2007Index = 2;
        private const int _2008_2013Index = 3;
        private const int _2014_2019Index = 4;


        public ExcelService(AssignmentDbContext dbContext)
        {
            this.dbContext = dbContext;
        }

        public async Task<byte[]> GetExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var excelFile = new ExcelPackage();
            using var excelWorkbook = excelFile.Workbook;

            using var populationSheet = excelWorkbook.Worksheets.Add("Population");

            populationSheet.OutLineSummaryBelow = false;
            populationSheet.Cells.Style.Font.Name = "Calibri";
            populationSheet.Cells.Style.Font.Size = 11;

            populationSheet.Cells["B1:D1"].Merge = true;
            populationSheet.Cells["B1"].Value = "Mean of the population";
            populationSheet.Cells[1, 1, 2, 1000].Style.Font.Bold = true;

            populationSheet.Cells["A2"].Value = "Country";
            populationSheet.Cells["B2"].Value = "2002-2007";
            populationSheet.Cells["C2"].Value = "2008-2013";
            populationSheet.Cells["D2"].Value = "2014-2019";

            var countriesData = dbContext.CountryData.AsNoTracking().AsEnumerable().GroupBy(c => c.Country).Where(c => c.Count() == 18).SelectMany(c => c).ToList();
            var populationMeanDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               Country = d.Key,
                                                               _2002_2007 = d.Average(e => e.Population)
                                                           })
                                                           .ToList();

            var tableRowIndex = 3;
            foreach (var countryData in populationMeanDataFor2002To2007)
            {
                populationSheet.Cells[$"A{tableRowIndex}"].Value = countryData.Country;
                populationSheet.Cells[$"B{tableRowIndex}"].Value = countryData._2002_2007;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            var populationMeanDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2008_2013 = d.Average(e => e.Population)
                                                           })
                                                           .ToList();

            foreach (var countryData in populationMeanDataFor2008To2013)
            {
                populationSheet.Cells[$"C{tableRowIndex}"].Value = countryData._2008_2013;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            var populationMeanDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2013_2019 = d.Average(e => e.Population)
                                                           })
                                                           .ToList();

            foreach (var countryData in populationMeanDataFor2014To2019)
            {
                populationSheet.Cells[$"D{tableRowIndex}"].Value = countryData._2013_2019;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            populationSheet.Cells["G1:I1"].Merge = true;
            populationSheet.Cells["G1"].Value = "Median of the population";

            populationSheet.Cells["F2"].Value = "Country";
            populationSheet.Cells["G2"].Value = "2002-2007";
            populationSheet.Cells["H2"].Value = "2008-2013";
            populationSheet.Cells["I2"].Value = "2014-2019";

            var populationMedianDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               Country = d.Key,
                                                               _2002_2007 = Median(d, e => e.Population)
                                                           })
                                                           .ToList();

            foreach (var countryData in populationMedianDataFor2002To2007)
            {
                populationSheet.Cells[$"F{tableRowIndex}"].Value = countryData.Country;
                populationSheet.Cells[$"G{tableRowIndex}"].Value = countryData._2002_2007;
                tableRowIndex++;
            }

            tableRowIndex = 3;
            var populationMedianDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2008_2013 = Median(d, e => e.Population)
                                                           })
                                                           .ToList();

            foreach (var countryData in populationMedianDataFor2008To2013)
            {
                populationSheet.Cells[$"H{tableRowIndex}"].Value = countryData._2008_2013;
                tableRowIndex++;
            }

            tableRowIndex = 3;
            var populationMedianDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2014_2019 = Median(d, e => e.Population)
                                                           })
                                                           .ToList();

            foreach (var countryData in populationMedianDataFor2014To2019)
            {
                populationSheet.Cells[$"I{tableRowIndex}"].Value = countryData._2014_2019;
                tableRowIndex++;
            }
            tableRowIndex = 3;

            populationSheet.Cells["L1:N1"].Merge = true;
            populationSheet.Cells["L1"].Value = "Mode of the population";

            populationSheet.Cells["K2"].Value = "Country";
            populationSheet.Cells["L2"].Value = "2002-2007";
            populationSheet.Cells["M2"].Value = "2008-2013";
            populationSheet.Cells["N2"].Value = "2014-2019";


            var populationModeDataForCountry = countriesData.GroupBy(d => d.Country).Select(d => d.Key).ToList();

            foreach (var countrydata in populationModeDataForCountry)
            {
                populationSheet.Cells[$"K{tableRowIndex}"].Value = countrydata;
                tableRowIndex++;
            }

            tableRowIndex = 3;

            var populationModeDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.Population)))
                                                               .ToList();

            foreach(var countryData in populationModeDataFor2002To2007)
            {
                if (countryData.HasValue)
                    populationSheet.Cells[$"L{tableRowIndex}"].Value = countryData;

                tableRowIndex++;
            }

            tableRowIndex = 3;

            var populationModeDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.Population)))
                                                               .ToList();

            foreach (var countryData in populationModeDataFor2008To2013)
            {
                if (countryData.HasValue)
                    populationSheet.Cells[$"M{tableRowIndex}"].Value = countryData;

                tableRowIndex++;
            }

            tableRowIndex = 3;

            var populationModeDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.Population)))
                                                               .ToList();

            foreach (var countryData in populationModeDataFor2014To2019)
            {
                if (countryData.HasValue)
                    populationSheet.Cells[$"N{tableRowIndex}"].Value = countryData;
                tableRowIndex++;
            }

            tableRowIndex = 3;

            populationSheet.Cells["Q1:S1"].Merge = true;
            populationSheet.Cells["Q1"].Value = "Standard deviation of the population";

            populationSheet.Cells["P2"].Value = "Country";
            populationSheet.Cells["Q2"].Value = "2002-2007";
            populationSheet.Cells["R2"].Value = "2008-2013";
            populationSheet.Cells["S2"].Value = "2014-2019";

            foreach(var countryData in populationModeDataForCountry)
            {
                populationSheet.Cells[$"P{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var populationSTDDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.Population - d.Average(f => f.Population), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach(var countryData in populationSTDDataFor2002To2007)
            {
                populationSheet.Cells[$"Q{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var populationSTDDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.Population - d.Average(f => f.Population), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in populationSTDDataFor2008To2013)
            {
                populationSheet.Cells[$"R{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var populationSTDDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.Population - d.Average(f => f.Population), 2))/(d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in populationSTDDataFor2014To2019)
            {
                populationSheet.Cells[$"S{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;

            populationSheet.Cells.AutoFitColumns();

            return excelFile.GetAsByteArray();
        }

        private static double? Median<TColl, TValue>(IEnumerable<TColl> source,
                                                    Func<TColl, TValue> selector)
        {
            return Median(source.Select<TColl, TValue>(selector));
        }

        private static double? Median<T>(IEnumerable<T> source)
        {
            if (Nullable.GetUnderlyingType(typeof(T)) != null)
                source = source.Where(x => x != null);

            int count = source.Count();
            if (count == 0)
                return null;

            source = source.OrderBy(n => n);

            int midpoint = count / 2;
            if (count % 2 == 0)
                return (Convert.ToDouble(source.ElementAt(midpoint - 1)) + Convert.ToDouble(source.ElementAt(midpoint))) / 2.0;
            else
                return Convert.ToDouble(source.ElementAt(midpoint));
        }

        private double? Mode(IEnumerable<double> collection)
        {
            var result =
                collection
                    .GroupBy(value => value)
                    .Where(value => value.Count() > 1);
            if (result.Any())
            {
                return result.OrderByDescending(group => group.Count())
                    .Select(group => group.Key)
                    .First();
            }

            return null;
        }

        public static double StdDev(IEnumerable<double> values)
        {
            double ret = 0;
            int count = values.Count();
            if (count > 1)
            {
                //Compute the Average
                double avg = values.Average();

                //Perform the Sum of (value-avg)^2
                double sum = values.Sum(d => (d - avg) * (d - avg));

                //Put it all together
                ret = Math.Sqrt((sum / count) - 1);
            }
            return ret;
        }
    }
}
