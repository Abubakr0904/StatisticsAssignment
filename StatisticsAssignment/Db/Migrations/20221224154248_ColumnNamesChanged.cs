using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace StatisticsAssignment.Db.Migrations
{
    /// <inheritdoc />
    public partial class ColumnNamesChanged : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Year",
                table: "CountryData",
                newName: "year");

            migrationBuilder.RenameColumn(
                name: "Country",
                table: "CountryData",
                newName: "country");

            migrationBuilder.RenameColumn(
                name: "ShareOfHouseholdConsumption",
                table: "CountryData",
                newName: "csh_c");

            migrationBuilder.RenameColumn(
                name: "RealGdp",
                table: "CountryData",
                newName: "rgdpna");

            migrationBuilder.RenameColumn(
                name: "Population",
                table: "CountryData",
                newName: "pop");

            migrationBuilder.RenameColumn(
                name: "AverageAnnualHoursWorkedByPersonsEngaged",
                table: "CountryData",
                newName: "avh");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "year",
                table: "CountryData",
                newName: "Year");

            migrationBuilder.RenameColumn(
                name: "country",
                table: "CountryData",
                newName: "Country");

            migrationBuilder.RenameColumn(
                name: "rgdpna",
                table: "CountryData",
                newName: "RealGdp");

            migrationBuilder.RenameColumn(
                name: "pop",
                table: "CountryData",
                newName: "Population");

            migrationBuilder.RenameColumn(
                name: "csh_c",
                table: "CountryData",
                newName: "ShareOfHouseholdConsumption");

            migrationBuilder.RenameColumn(
                name: "avh",
                table: "CountryData",
                newName: "AverageAnnualHoursWorkedByPersonsEngaged");
        }
    }
}
