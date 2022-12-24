namespace StatisticsAssignment.Db.Entities
{
    public class CountryData
    {
        public virtual int Id { get; set; }

        public string Country { get; set; }

        public int Year { get; set; }

        public decimal Population { get; set; }

        public decimal AverageAnnualHoursWorkedByPersonsEngaged { get; set; }

        public decimal RealGdp { get; set; }

        public decimal ShareOfHouseholdConsumption { get; set; }
    }
}
