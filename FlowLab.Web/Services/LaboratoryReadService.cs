using Microsoft.Data.Sqlite;

namespace FlowLab.Web.Services;

public sealed class LaboratoryReadService(IWebHostEnvironment environment)
{
    private readonly string _connectionString = new SqliteConnectionStringBuilder { DataSource = Path.GetFullPath(Path.Combine(environment.ContentRootPath, "..", "data", "flowlab.db")) }.ToString();

    public DashboardSummary GetDashboardSummary()
    {
        using var connection = new SqliteConnection(_connectionString); connection.Open();
        return new DashboardSummary(Count(connection, "SELECT COUNT(*) FROM laboratory_equipment"), Count(connection, "SELECT COUNT(*) FROM laboratory_equipment WHERE include_in_calibration_programme = 1"), Count(connection, "SELECT COUNT(*) FROM controlled_documents"), Count(connection, "SELECT COUNT(*) FROM measurement_sessions"));
    }

    private static int Count(SqliteConnection connection, string sql)
    {
        using var command = connection.CreateCommand(); command.CommandText = sql;
        return Convert.ToInt32(command.ExecuteScalar());
    }
}

public record DashboardSummary(int Equipment, int Programme, int Documents, int Sessions);
