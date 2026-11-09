using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace MatOffice.Data.Migrations
{
    /// <inheritdoc />
    public partial class AppModelBindings : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "ConsumerId",
                table: "HostClient",
                type: "character varying(64)",
                maxLength: 64,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "ManagedBy",
                table: "HostClient",
                type: "character varying(40)",
                maxLength: 40,
                nullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_HostClient_ConsumerId",
                table: "HostClient",
                column: "ConsumerId",
                unique: true,
                filter: "\"ConsumerId\" IS NOT NULL AND \"UpdateState\" <> 0");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_HostClient_ConsumerId",
                table: "HostClient");

            migrationBuilder.DropColumn(
                name: "ConsumerId",
                table: "HostClient");

            migrationBuilder.DropColumn(
                name: "ManagedBy",
                table: "HostClient");
        }
    }
}
