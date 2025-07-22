using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApplicationAPI.Migrations
{
    /// <inheritdoc />
    public partial class newMig : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "FitToConceptTestTable",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProjectName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TestName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TestNameColor = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Average = table.Column<double>(type: "float", nullable: false),
                    TestNameBold = table.Column<bool>(type: "bit", nullable: false),
                    AverageColor = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    ProjectTemplateType = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_FitToConceptTestTable", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "OverallImpressionsTestTable",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ProjectName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TestName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Positive = table.Column<double>(type: "float", nullable: false),
                    Neutral = table.Column<double>(type: "float", nullable: false),
                    Negative = table.Column<double>(type: "float", nullable: false),
                    TestNameBold = table.Column<bool>(type: "bit", nullable: true),
                    TestNameColor = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    ProjectTemplateType = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_OverallImpressionsTestTable", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "FitToConceptTestTable");

            migrationBuilder.DropTable(
                name: "OverallImpressionsTestTable");
        }
    }
}
