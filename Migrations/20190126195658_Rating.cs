using Microsoft.EntityFrameworkCore.Migrations;

namespace RazorPagesHolland.Migrations
{
    public partial class Rating : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "MapUrl",
                table: "Holland");

            migrationBuilder.AddColumn<float>(
                name: "Latitude",
                table: "Holland",
                type: "float(10, 6)",
                nullable: false,
                defaultValue: 0f);

            migrationBuilder.AddColumn<float>(
                name: "Longitude",
                table: "Holland",
                type: "float(10, 6)",
                nullable: false,
                defaultValue: 0f);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Latitude",
                table: "Holland");

            migrationBuilder.DropColumn(
                name: "Longitude",
                table: "Holland");

            migrationBuilder.AddColumn<string>(
                name: "MapUrl",
                table: "Holland",
                nullable: true);
        }
    }
}
