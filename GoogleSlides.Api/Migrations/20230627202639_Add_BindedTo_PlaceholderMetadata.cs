using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace GoogleSlides.Api.Migrations
{
    public partial class Add_BindedTo_PlaceholderMetadata : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "BindedTo",
                table: "Placeholders",
                type: "TEXT",
                nullable: false,
                defaultValue: "");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "BindedTo",
                table: "Placeholders");
        }
    }
}
