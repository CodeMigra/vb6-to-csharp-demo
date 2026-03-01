using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodeMigra.Inventory
{
    public class InventoryItem
    {
        public long ItemId { get; init; }
        public string SKU { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public int QuantityOnHand { get; set; }
        public int ReorderLevel { get; set; }
    }

    public interface IInventoryRepository
    {
        Task<IEnumerable<InventoryItem>> GetAllAsync();
        Task<InventoryItem?> GetByIdAsync(long id);
        Task<long> InsertAsync(InventoryItem item);
        Task UpdateAsync(InventoryItem item);
        Task DeleteAsync(long id);
    }

    public class SqliteInventoryRepository : IInventoryRepository, IDisposable
    {
        private readonly SQLiteConnection _connection;

        public SqliteInventoryRepository(string dbPath)
        {
            _connection = new SQLiteConnection($"Data Source={dbPath};Version=3;");
            _connection.Open();
            EnsureSchema();
        }

        private void EnsureSchema()
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS Items (
                    ItemId       INTEGER PRIMARY KEY AUTOINCREMENT,
                    SKU          TEXT NOT NULL UNIQUE,
                    Description  TEXT,
                    Price        TEXT NOT NULL DEFAULT '0.000000',
                    QuantityOnHand INTEGER NOT NULL DEFAULT 0,
                    ReorderLevel   INTEGER NOT NULL DEFAULT 0
                )";
            cmd.ExecuteNonQuery();
        }

        public async Task<IEnumerable<InventoryItem>> GetAllAsync()
        {
            var items = new List<InventoryItem>();
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT * FROM Items ORDER BY SKU";

            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
                items.Add(MapRow(reader));

            return items;
        }

        public async Task<InventoryItem?> GetByIdAsync(long id)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "SELECT * FROM Items WHERE ItemId = @id";
            cmd.Parameters.AddWithValue("@id", id);

            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapRow(reader) : null;
        }

        public async Task<long> InsertAsync(InventoryItem item)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO Items (SKU, Description, Price, QuantityOnHand, ReorderLevel)
                VALUES (@sku, @desc, @price, @qty, @reorder);
                SELECT last_insert_rowid();";
            AddItemParams(cmd, item);
            return (long)(await cmd.ExecuteScalarAsync())!;
        }

        public async Task UpdateAsync(InventoryItem item)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = @"
                UPDATE Items SET
                    SKU = @sku, Description = @desc, Price = @price,
                    QuantityOnHand = @qty, ReorderLevel = @reorder
                WHERE ItemId = @id";
            AddItemParams(cmd, item);
            cmd.Parameters.AddWithValue("@id", item.ItemId);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task DeleteAsync(long id)
        {
            using var cmd = _connection.CreateCommand();
            cmd.CommandText = "DELETE FROM Items WHERE ItemId = @id";
            cmd.Parameters.AddWithValue("@id", id);
            await cmd.ExecuteNonQueryAsync();
        }

        private static void AddItemParams(SQLiteCommand cmd, InventoryItem item)
        {
            cmd.Parameters.AddWithValue("@sku",    item.SKU);
            cmd.Parameters.AddWithValue("@desc",   item.Description);
            // Store price as TEXT to avoid double precision loss (SQLite REAL is IEEE 754 64-bit)
            cmd.Parameters.AddWithValue("@price",  item.Price.ToString("F6", System.Globalization.CultureInfo.InvariantCulture));
            cmd.Parameters.AddWithValue("@qty",    item.QuantityOnHand);
            cmd.Parameters.AddWithValue("@reorder",item.ReorderLevel);
        }

        private static InventoryItem MapRow(IDataRecord r) => new()
        {
            ItemId         = r.GetInt64(r.GetOrdinal("ItemId")),
            SKU            = r.GetString(r.GetOrdinal("SKU")),
            Description    = r.IsDBNull(r.GetOrdinal("Description")) ? "" : r.GetString(r.GetOrdinal("Description")),
            Price          = decimal.Parse(r.GetString(r.GetOrdinal("Price")), System.Globalization.CultureInfo.InvariantCulture),
            QuantityOnHand = r.GetInt32(r.GetOrdinal("QuantityOnHand")),
            ReorderLevel   = r.GetInt32(r.GetOrdinal("ReorderLevel"))
        };

        public void Dispose() => _connection.Dispose();
    }
}
