using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace StatisticsAssignment.Db.Entities
{
    public class CountryData
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public virtual int Id { get; set; }

        [Required, Column("country")]
        public string Country { get; set; }

        [Required, Column("year")]
        public int Year { get; set; }

        [Required, Column("pop")]
        public double Population { get; set; }

        [Required, Column("avh")]
        public double AverageAnnualHoursWorkedByPersonsEngaged { get; set; }

        [Required, Column("rgdpna")]
        public double RealGdp { get; set; }

        [Required, Column("csh_c")]
        public double ShareOfHouseholdConsumption { get; set; }
    }
}
