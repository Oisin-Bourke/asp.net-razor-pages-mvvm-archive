using System;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;

namespace RazorPagesHolland.Migrations
{
    public partial class Initial : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "VideoUrl",
                table: "Holland",
                nullable: false,
                oldClrType: typeof(string),
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "VesselName",
                table: "Holland",
                nullable: false,
                oldClrType: typeof(string),
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "ROVDiveName",
                table: "Holland",
                nullable: false,
                oldClrType: typeof(string),
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "Location",
                table: "Holland",
                nullable: false,
                oldClrType: typeof(string),
                oldNullable: true);

            migrationBuilder.CreateTable(
                name: "Survey",
                columns: table => new
                {
                    SurveyID = table.Column<int>(nullable: false)
                        .Annotation("MySql:ValueGenerationStrategy", MySqlValueGenerationStrategy.IdentityColumn),
                    TimeInterval = table.Column<DateTime>(nullable: false),
                    VesselNumber = table.Column<int>(nullable: false),
                    Latitude = table.Column<float>(type: "float(10, 6)", nullable: false),
                    Longitude = table.Column<float>(type: "float(10, 6)", nullable: false),
                    HollandID = table.Column<int>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Survey", x => x.SurveyID);
                    table.ForeignKey(
                        name: "FK_Survey_Holland_HollandID",
                        column: x => x.HollandID,
                        principalTable: "Holland",
                        principalColumn: "ID",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Survey_HollandID",
                table: "Survey",
                column: "HollandID");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Survey");

            migrationBuilder.AlterColumn<string>(
                name: "VideoUrl",
                table: "Holland",
                nullable: true,
                oldClrType: typeof(string));

            migrationBuilder.AlterColumn<string>(
                name: "VesselName",
                table: "Holland",
                nullable: true,
                oldClrType: typeof(string));

            migrationBuilder.AlterColumn<string>(
                name: "ROVDiveName",
                table: "Holland",
                nullable: true,
                oldClrType: typeof(string));

            migrationBuilder.AlterColumn<string>(
                name: "Location",
                table: "Holland",
                nullable: true,
                oldClrType: typeof(string));
        }
    }
}
