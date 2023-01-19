using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using StatisticsAssignment.Db;

namespace StatisticsAssignment
{
    public sealed class ExcelService
    {
        private readonly AssignmentDbContext dbContext;

        public ExcelService(AssignmentDbContext dbContext)
        {
            this.dbContext = dbContext;
        }

        public async Task<byte[]> GetExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var excelFile = new ExcelPackage();
            using var excelWorkbook = excelFile.Workbook;

            using var populationSheet = excelWorkbook.Worksheets.Add("POP");

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

            using var gdpSheet = excelWorkbook.Worksheets.Add("GDP");

            gdpSheet.OutLineSummaryBelow = false;
            gdpSheet.Cells.Style.Font.Name = "Calibri";
            gdpSheet.Cells.Style.Font.Size = 11;

            gdpSheet.Cells["B1:D1"].Merge = true;
            gdpSheet.Cells["B1"].Value = "Mean of the GDP";
            gdpSheet.Cells[1, 1, 2, 1000].Style.Font.Bold = true;

            gdpSheet.Cells["A2"].Value = "Country";
            gdpSheet.Cells["B2"].Value = "2002-2007";
            gdpSheet.Cells["C2"].Value = "2008-2013";
            gdpSheet.Cells["D2"].Value = "2014-2019";

            var gdpMeanDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               Country = d.Key,
                                                               _2002_2007 = d.Average(e => e.RealGdp)
                                                           })
                                                           .ToList();

            tableRowIndex = 3;
            foreach (var countryData in gdpMeanDataFor2002To2007)
            {
                gdpSheet.Cells[$"A{tableRowIndex}"].Value = countryData.Country;
                gdpSheet.Cells[$"B{tableRowIndex}"].Value = countryData._2002_2007;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            var gdpMeanDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2008_2013 = d.Average(e => e.RealGdp)
                                                           })
                                                           .ToList();

            foreach (var countryData in gdpMeanDataFor2008To2013)
            {
                gdpSheet.Cells[$"C{tableRowIndex}"].Value = countryData._2008_2013;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            var gdpMeanDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2013_2019 = d.Average(e => e.RealGdp)
                                                           })
                                                           .ToList();

            foreach (var countryData in gdpMeanDataFor2014To2019)
            {
                gdpSheet.Cells[$"D{tableRowIndex}"].Value = countryData._2013_2019;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            gdpSheet.Cells["G1:I1"].Merge = true;
            gdpSheet.Cells["G1"].Value = "Median of the GDP";

            gdpSheet.Cells["F2"].Value = "Country";
            gdpSheet.Cells["G2"].Value = "2002-2007";
            gdpSheet.Cells["H2"].Value = "2008-2013";
            gdpSheet.Cells["I2"].Value = "2014-2019";

            var gdpMedianDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               Country = d.Key,
                                                               _2002_2007 = Median(d, e => e.RealGdp)
                                                           })
                                                           .ToList();

            foreach (var countryData in gdpMedianDataFor2002To2007)
            {
                gdpSheet.Cells[$"F{tableRowIndex}"].Value = countryData.Country;
                gdpSheet.Cells[$"G{tableRowIndex}"].Value = countryData._2002_2007;
                tableRowIndex++;
            }

            tableRowIndex = 3;
            var gdpMedianDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2008_2013 = Median(d, e => e.RealGdp)
                                                           })
                                                           .ToList();

            foreach (var countryData in gdpMedianDataFor2008To2013)
            {
                gdpSheet.Cells[$"H{tableRowIndex}"].Value = countryData._2008_2013;
                tableRowIndex++;
            }

            tableRowIndex = 3;
            var gdpMedianDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2014_2019 = Median(d, e => e.RealGdp)
                                                           })
                                                           .ToList();

            foreach (var countryData in gdpMedianDataFor2014To2019)
            {
                gdpSheet.Cells[$"I{tableRowIndex}"].Value = countryData._2014_2019;
                tableRowIndex++;
            }
            tableRowIndex = 3;

            gdpSheet.Cells["L1:N1"].Merge = true;
            gdpSheet.Cells["L1"].Value = "Mode of the GDP";

            gdpSheet.Cells["K2"].Value = "Country";
            gdpSheet.Cells["L2"].Value = "2002-2007";
            gdpSheet.Cells["M2"].Value = "2008-2013";
            gdpSheet.Cells["N2"].Value = "2014-2019";


            var gdpModeDataForCountry = countriesData.GroupBy(d => d.Country).Select(d => d.Key).ToList();

            foreach (var countrydata in gdpModeDataForCountry)
            {
                gdpSheet.Cells[$"K{tableRowIndex}"].Value = countrydata;
                tableRowIndex++;
            }

            tableRowIndex = 3;

            var gdpModeDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.RealGdp)))
                                                               .ToList();

            foreach (var countryData in gdpModeDataFor2002To2007)
            {
                if (countryData.HasValue)
                    gdpSheet.Cells[$"L{tableRowIndex}"].Value = countryData;

                tableRowIndex++;
            }

            tableRowIndex = 3;

            var gdpModeDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.RealGdp)))
                                                               .ToList();

            foreach (var countryData in gdpModeDataFor2008To2013)
            {
                if (countryData.HasValue)
                    gdpSheet.Cells[$"M{tableRowIndex}"].Value = countryData;

                tableRowIndex++;
            }

            tableRowIndex = 3;

            var gdpModeDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.RealGdp)))
                                                               .ToList();

            foreach (var countryData in gdpModeDataFor2014To2019)
            {
                if (countryData.HasValue)
                    gdpSheet.Cells[$"N{tableRowIndex}"].Value = countryData;
                tableRowIndex++;
            }

            tableRowIndex = 3;

            gdpSheet.Cells["Q1:S1"].Merge = true;
            gdpSheet.Cells["Q1"].Value = "Standard deviation of the GDP";

            gdpSheet.Cells["P2"].Value = "Country";
            gdpSheet.Cells["Q2"].Value = "2002-2007";
            gdpSheet.Cells["R2"].Value = "2008-2013";
            gdpSheet.Cells["S2"].Value = "2014-2019";

            foreach (var countryData in populationModeDataForCountry)
            {
                gdpSheet.Cells[$"P{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var gdpSTDDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.RealGdp - d.Average(f => f.RealGdp), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in gdpSTDDataFor2002To2007)
            {
                gdpSheet.Cells[$"Q{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var gdpSTDDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.RealGdp - d.Average(f => f.RealGdp), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in gdpSTDDataFor2008To2013)
            {
                gdpSheet.Cells[$"R{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var gdpSTDDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.RealGdp - d.Average(f => f.RealGdp), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in gdpSTDDataFor2014To2019)
            {
                gdpSheet.Cells[$"S{tableRowIndex++}"].Value = countryData;
            }

            using var AVHSheet = excelWorkbook.Worksheets.Add("AVH");

            AVHSheet.OutLineSummaryBelow = false;
            AVHSheet.Cells.Style.Font.Name = "Calibri";
            AVHSheet.Cells.Style.Font.Size = 11;

            AVHSheet.Cells["B1:D1"].Merge = true;
            AVHSheet.Cells["B1"].Value = "Mean of the Average Annual Hours Worked";
            AVHSheet.Cells[1, 1, 2, 1000].Style.Font.Bold = true;

            AVHSheet.Cells["A2"].Value = "Country";
            AVHSheet.Cells["B2"].Value = "2002-2007";
            AVHSheet.Cells["C2"].Value = "2008-2013";
            AVHSheet.Cells["D2"].Value = "2014-2019";

            var avhMeanDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               Country = d.Key,
                                                               _2002_2007 = d.Average(e => e.AverageAnnualHoursWorkedByPersonsEngaged)
                                                           })
                                                           .ToList();

            tableRowIndex = 3;
            foreach (var countryData in avhMeanDataFor2002To2007)
            {
                AVHSheet.Cells[$"A{tableRowIndex}"].Value = countryData.Country;
                AVHSheet.Cells[$"B{tableRowIndex}"].Value = countryData._2002_2007;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            var avhMeanDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2008_2013 = d.Average(e => e.AverageAnnualHoursWorkedByPersonsEngaged)
                                                           })
                                                           .ToList();

            foreach (var countryData in avhMeanDataFor2008To2013)
            {
                AVHSheet.Cells[$"C{tableRowIndex}"].Value = countryData._2008_2013;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            var avhMeanDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2013_2019 = d.Average(e => e.AverageAnnualHoursWorkedByPersonsEngaged)
                                                           })
                                                           .ToList();

            foreach (var countryData in avhMeanDataFor2014To2019)
            {
                AVHSheet.Cells[$"D{tableRowIndex}"].Value = countryData._2013_2019;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            AVHSheet.Cells["G1:I1"].Merge = true;
            AVHSheet.Cells["G1"].Value = "Median of the Average Annual Hours Worked";

            AVHSheet.Cells["F2"].Value = "Country";
            AVHSheet.Cells["G2"].Value = "2002-2007";
            AVHSheet.Cells["H2"].Value = "2008-2013";
            AVHSheet.Cells["I2"].Value = "2014-2019";

            var avhMedianDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               Country = d.Key,
                                                               _2002_2007 = Median(d, e => e.AverageAnnualHoursWorkedByPersonsEngaged)
                                                           })
                                                           .ToList();

            foreach (var countryData in avhMedianDataFor2002To2007)
            {
                AVHSheet.Cells[$"F{tableRowIndex}"].Value = countryData.Country;
                AVHSheet.Cells[$"G{tableRowIndex}"].Value = countryData._2002_2007;
                tableRowIndex++;
            }

            tableRowIndex = 3;
            var avhMedianDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2008_2013 = Median(d, e => e.AverageAnnualHoursWorkedByPersonsEngaged)
                                                           })
                                                           .ToList();

            foreach (var countryData in avhMedianDataFor2008To2013)
            {
                AVHSheet.Cells[$"H{tableRowIndex}"].Value = countryData._2008_2013;
                tableRowIndex++;
            }

            tableRowIndex = 3;
            var avhMedianDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2014_2019 = Median(d, e => e.AverageAnnualHoursWorkedByPersonsEngaged)
                                                           })
                                                           .ToList();

            foreach (var countryData in avhMedianDataFor2014To2019)
            {
                AVHSheet.Cells[$"I{tableRowIndex}"].Value = countryData._2014_2019;
                tableRowIndex++;
            }
            tableRowIndex = 3;

            AVHSheet.Cells["L1:N1"].Merge = true;
            AVHSheet.Cells["L1"].Value = "Mode of the Average Annual Hours Worked";

            AVHSheet.Cells["K2"].Value = "Country";
            AVHSheet.Cells["L2"].Value = "2002-2007";
            AVHSheet.Cells["M2"].Value = "2008-2013";
            AVHSheet.Cells["N2"].Value = "2014-2019";


            var avhModeDataForCountry = countriesData.GroupBy(d => d.Country).Select(d => d.Key).ToList();

            foreach (var countrydata in avhModeDataForCountry)
            {
                AVHSheet.Cells[$"K{tableRowIndex}"].Value = countrydata;
                tableRowIndex++;
            }

            tableRowIndex = 3;

            var avhModeDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.AverageAnnualHoursWorkedByPersonsEngaged)))
                                                               .ToList();

            foreach (var countryData in avhModeDataFor2002To2007)
            {
                if (countryData.HasValue)
                    AVHSheet.Cells[$"L{tableRowIndex}"].Value = countryData;

                tableRowIndex++;
            }

            tableRowIndex = 3;

            var avhModeDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.AverageAnnualHoursWorkedByPersonsEngaged)))
                                                               .ToList();

            foreach (var countryData in avhModeDataFor2008To2013)
            {
                if (countryData.HasValue)
                    AVHSheet.Cells[$"M{tableRowIndex}"].Value = countryData;

                tableRowIndex++;
            }

            tableRowIndex = 3;

            var avhModeDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.AverageAnnualHoursWorkedByPersonsEngaged)))
                                                               .ToList();

            foreach (var countryData in avhModeDataFor2014To2019)
            {
                if (countryData.HasValue)
                    AVHSheet.Cells[$"N{tableRowIndex}"].Value = countryData;
                tableRowIndex++;
            }

            tableRowIndex = 3;

            AVHSheet.Cells["Q1:S1"].Merge = true;
            AVHSheet.Cells["Q1"].Value = "Standard deviation of the Average Annual Hours Worked";

            AVHSheet.Cells["P2"].Value = "Country";
            AVHSheet.Cells["Q2"].Value = "2002-2007";
            AVHSheet.Cells["R2"].Value = "2008-2013";
            AVHSheet.Cells["S2"].Value = "2014-2019";

            foreach (var countryData in avhModeDataForCountry)
            {
                AVHSheet.Cells[$"P{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var avhSTDDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.AverageAnnualHoursWorkedByPersonsEngaged - d.Average(f => f.AverageAnnualHoursWorkedByPersonsEngaged), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in avhSTDDataFor2002To2007)
            {
                AVHSheet.Cells[$"Q{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var avhSTDDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.AverageAnnualHoursWorkedByPersonsEngaged - d.Average(f => f.AverageAnnualHoursWorkedByPersonsEngaged), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in avhSTDDataFor2008To2013)
            {
                AVHSheet.Cells[$"R{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var avhSTDDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.AverageAnnualHoursWorkedByPersonsEngaged - d.Average(f => f.AverageAnnualHoursWorkedByPersonsEngaged), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in avhSTDDataFor2014To2019)
            {
                AVHSheet.Cells[$"S{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;

            using var cshcSheet = excelWorkbook.Worksheets.Add("CSH_C");

            cshcSheet.OutLineSummaryBelow = false;
            cshcSheet.Cells.Style.Font.Name = "Calibri";
            cshcSheet.Cells.Style.Font.Size = 11;

            cshcSheet.Cells["B1:D1"].Merge = true;
            cshcSheet.Cells["B1"].Value = "Mean of the Share of Household Consumption at Current PPPs";
            cshcSheet.Cells[1, 1, 2, 1000].Style.Font.Bold = true;

            cshcSheet.Cells["A2"].Value = "Country";
            cshcSheet.Cells["B2"].Value = "2002-2007";
            cshcSheet.Cells["C2"].Value = "2008-2013";
            cshcSheet.Cells["D2"].Value = "2014-2019";

            var cshcMeanDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               Country = d.Key,
                                                               _2002_2007 = d.Average(e => e.ShareOfHouseholdConsumption)
                                                           })
                                                           .ToList();

            foreach (var countryData in cshcMeanDataFor2002To2007)
            {
                cshcSheet.Cells[$"A{tableRowIndex}"].Value = countryData.Country;
                cshcSheet.Cells[$"B{tableRowIndex}"].Value = countryData._2002_2007;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            var cshcMeanDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2008_2013 = d.Average(e => e.ShareOfHouseholdConsumption)
                                                           })
                                                           .ToList();

            foreach (var countryData in cshcMeanDataFor2008To2013)
            {
                cshcSheet.Cells[$"C{tableRowIndex}"].Value = countryData._2008_2013;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            var cshcMeanDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2013_2019 = d.Average(e => e.ShareOfHouseholdConsumption)
                                                           })
                                                           .ToList();

            foreach (var countryData in cshcMeanDataFor2014To2019)
            {
                cshcSheet.Cells[$"D{tableRowIndex}"].Value = countryData._2013_2019;

                tableRowIndex++;
            }

            tableRowIndex = 3;
            cshcSheet.Cells["G1:I1"].Merge = true;
            cshcSheet.Cells["G1"].Value = "Median of the Share of Household Consumption at Current PPPs";

            cshcSheet.Cells["F2"].Value = "Country";
            cshcSheet.Cells["G2"].Value = "2002-2007";
            cshcSheet.Cells["H2"].Value = "2008-2013";
            cshcSheet.Cells["I2"].Value = "2014-2019";

            var cshcMedianDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               Country = d.Key,
                                                               _2002_2007 = Median(d, e => e.ShareOfHouseholdConsumption)
                                                           })
                                                           .ToList();

            foreach (var countryData in cshcMedianDataFor2002To2007)
            {
                cshcSheet.Cells[$"F{tableRowIndex}"].Value = countryData.Country;
                cshcSheet.Cells[$"G{tableRowIndex}"].Value = countryData._2002_2007;
                tableRowIndex++;
            }

            tableRowIndex = 3;
            var cshcMedianDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2008_2013 = Median(d, e => e.ShareOfHouseholdConsumption)
                                                           })
                                                           .ToList();

            foreach (var countryData in cshcMedianDataFor2008To2013)
            {
                cshcSheet.Cells[$"H{tableRowIndex}"].Value = countryData._2008_2013;
                tableRowIndex++;
            }

            tableRowIndex = 3;
            var cshcMedianDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                           .GroupBy(d => d.Country)
                                                           .Select(d => new
                                                           {
                                                               _2014_2019 = Median(d, e => e.ShareOfHouseholdConsumption)
                                                           })
                                                           .ToList();

            foreach (var countryData in cshcMedianDataFor2014To2019)
            {
                cshcSheet.Cells[$"I{tableRowIndex}"].Value = countryData._2014_2019;
                tableRowIndex++;
            }
            tableRowIndex = 3;

            cshcSheet.Cells["L1:N1"].Merge = true;
            cshcSheet.Cells["L1"].Value = "Mode of the Share of Household Consumption at Current PPPs";

            cshcSheet.Cells["K2"].Value = "Country";
            cshcSheet.Cells["L2"].Value = "2002-2007";
            cshcSheet.Cells["M2"].Value = "2008-2013";
            cshcSheet.Cells["N2"].Value = "2014-2019";


            var cshcModeDataForCountry = countriesData.GroupBy(d => d.Country).Select(d => d.Key).ToList();

            foreach (var countrydata in cshcModeDataForCountry)
            {
                cshcSheet.Cells[$"K{tableRowIndex}"].Value = countrydata;
                tableRowIndex++;
            }

            tableRowIndex = 3;

            var cshcModeDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.ShareOfHouseholdConsumption)))
                                                               .ToList();

            foreach (var countryData in cshcModeDataFor2002To2007)
            {
                if (countryData.HasValue)
                    cshcSheet.Cells[$"L{tableRowIndex}"].Value = countryData;

                tableRowIndex++;
            }

            tableRowIndex = 3;

            var cshcModeDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.ShareOfHouseholdConsumption)))
                                                               .ToList();

            foreach (var countryData in cshcModeDataFor2008To2013)
            {
                if (countryData.HasValue)
                    cshcSheet.Cells[$"M{tableRowIndex}"].Value = countryData;

                tableRowIndex++;
            }

            tableRowIndex = 3;

            var cshcModeDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                               .GroupBy(d => d.Country)
                                                               .Select(d => Mode(d.Select(e => e.ShareOfHouseholdConsumption)))
                                                               .ToList();

            foreach (var countryData in cshcModeDataFor2014To2019)
            {
                if (countryData.HasValue)
                    cshcSheet.Cells[$"N{tableRowIndex}"].Value = countryData;
                tableRowIndex++;
            }

            tableRowIndex = 3;

            cshcSheet.Cells["Q1:S1"].Merge = true;
            cshcSheet.Cells["Q1"].Value = "Standard deviation of the Share of Household Consumption at Current PPPs";

            cshcSheet.Cells["P2"].Value = "Country";
            cshcSheet.Cells["Q2"].Value = "2002-2007";
            cshcSheet.Cells["R2"].Value = "2008-2013";
            cshcSheet.Cells["S2"].Value = "2014-2019";

            foreach (var countryData in avhModeDataForCountry)
            {
                cshcSheet.Cells[$"P{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var cshcSTDDataFor2002To2007 = countriesData.Where(d => d.Year >= 2002 && d.Year <= 2007)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.ShareOfHouseholdConsumption - d.Average(f => f.ShareOfHouseholdConsumption), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in cshcSTDDataFor2002To2007)
            {
                cshcSheet.Cells[$"Q{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var cshcSTDDataFor2008To2013 = countriesData.Where(d => d.Year >= 2008 && d.Year <= 2013)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.ShareOfHouseholdConsumption - d.Average(f => f.ShareOfHouseholdConsumption), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in cshcSTDDataFor2008To2013)
            {
                cshcSheet.Cells[$"R{tableRowIndex++}"].Value = countryData;
            }

            tableRowIndex = 3;
            var cshcSTDDataFor2014To2019 = countriesData.Where(d => d.Year >= 2014 && d.Year <= 2019)
                                                              .GroupBy(d => d.Country)
                                                              .Select(d => Math.Sqrt((d.Sum(e => Math.Pow(e.ShareOfHouseholdConsumption - d.Average(f => f.ShareOfHouseholdConsumption), 2)) / (d.Count() - 1))))
                                                              .ToList();

            foreach (var countryData in cshcSTDDataFor2014To2019)
            {
                cshcSheet.Cells[$"S{tableRowIndex++}"].Value = countryData;
            }

            using var gdpGrowthRateSheet = excelWorkbook.Worksheets.Add("Gdp_Growth_Rate");
            gdpGrowthRateSheet.Cells[1, 1].Value = "Gdp Growth Rate for Countries by Years";
            int startYear = 2002;
            int i;
            for (i = 2; i <= 18; i++)
            {
                gdpGrowthRateSheet.Cells[1, i].Value = $"{startYear + i - 2}-{startYear + i - 1}";
            }
            gdpGrowthRateSheet.Cells[1, i].Value = "Mean Gdp Growth rate";


            var gdpCountriesData = await dbContext.CountryData.AsNoTracking()
                                                  .GroupBy(c => c.Country)
                                                  .Select(c => new
                                                  {
                                                      Country = c.Key,
                                                      YearsData = c.ToList()
                                                  })
                                                  .ToArrayAsync();
            tableRowIndex = 2;

            foreach (var countryData in gdpCountriesData)
            {
                gdpGrowthRateSheet.Cells[tableRowIndex, 1].Value = countryData.Country;
                var growthRates = new List<double?>();
                int columnIndex;
                for (columnIndex = 2; columnIndex <= 18; columnIndex++)
                {
                    var gdpGrowthRate = countryData.YearsData.Find(d => d.Year == startYear + columnIndex - 1)?.RealGdp / countryData.YearsData.Find(d => d.Year == startYear + columnIndex - 2)?.RealGdp;
                    growthRates.Add(gdpGrowthRate);
                    gdpGrowthRateSheet.Cells[tableRowIndex, columnIndex].Value = gdpGrowthRate;
                }
                gdpGrowthRateSheet.Cells[tableRowIndex, columnIndex].Value = growthRates.Average(r => r);
                tableRowIndex++;
            }

            populationSheet.Cells.AutoFitColumns();
            gdpSheet.Cells.AutoFitColumns();
            AVHSheet.Cells.AutoFitColumns();
            cshcSheet.Cells.AutoFitColumns();
            gdpGrowthRateSheet.Cells.AutoFitColumns();

            return excelFile.GetAsByteArray();
        }

        private static double? Median<TColl, TValue>(IEnumerable<TColl> source,
                                                    Func<TColl, TValue> selector)
        {
            return Median(source.Select(selector));
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
