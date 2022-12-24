using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace StatisticsAssignment.Db.Migrations
{
    /// <inheritdoc />
    public partial class CountryDataEntityAdded : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "CountryData",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Country = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Year = table.Column<int>(type: "int", nullable: false),
                    Population = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    AverageAnnualHoursWorkedByPersonsEngaged = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    RealGdp = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    ShareOfHouseholdConsumption = table.Column<decimal>(type: "decimal(18,2)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_CountryData", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "CountryData");
        }
    }
}
