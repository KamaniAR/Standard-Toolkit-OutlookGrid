using System.Globalization;

using Krypton.Toolkit;

namespace OutlookGridTest;

public class ProductDto
{
    public int Id { get; set; }
    public string Name { get; set; }
    public string Category { get; set; }
    public decimal Price { get; set; } // Using decimal for currency
    public int StockQuantity { get; set; }
    public DateTime LastRestockDate { get; set; }
    public bool IsAvailable { get; set; }
    public byte Rating { get; set; } // For a double type
    public Image? Thumbnail { get; set; } // For an Image type (needs System.Drawing)
    public double Progress { get; set; }

    // Constructor for easy initialization
    public ProductDto(int id, string name, string category, decimal price, int stockQuantity, DateTime lastRestockDate, bool isAvailable, byte rating, Image? thumbnail = null, double progress = 0)
    {
        Id = id;
        Name = name;
        Category = category;
        Price = price;
        StockQuantity = stockQuantity;
        LastRestockDate = lastRestockDate;
        IsAvailable = isAvailable;
        Rating = rating;
        Thumbnail = thumbnail;
        Progress = progress;
    }
}


public static class GridHelper
{
    public static void CreateOutlookGridColumn(this KryptonOutlookGrid grid, string columnName, string displayName, int width, int displayIndex, bool visible = true,
        bool showTotal = false, SortOrder sortOrder = SortOrder.None, int groupIndex = -1, DataGridViewContentAlignment alignment = DataGridViewContentAlignment.MiddleLeft,
        int decimalPlace = -1, KryptonOutlookGridAggregationType aggregationType = KryptonOutlookGridAggregationType.None, string dataType = "string")
    {
        try
        {
            DataGridViewColumn dataGridViewColumn = CreateGridColumn(columnName, displayName, width, displayIndex, visible, showTotal ? "S" : "", dataType, decimalPlace, (int)alignment);
            grid.Columns.Insert(displayIndex, dataGridViewColumn);
            grid.AddInternalColumn(dataGridViewColumn, new OutlookGridDefaultGroup(null), sortOrder, groupIndex, (sortOrder == SortOrder.None) ? (-1) : 0, aggregationType);
        }
        catch (Exception)
        {
            throw;
        }
    }

    private static DataGridViewColumn CreateGridColumn(string columnName, string displayName, int width, int displayIndex, bool visible, string tag, string dataType, int decimalPlace, int alignment)
    {
        DataGridViewTextBoxColumn obj = new DataGridViewTextBoxColumn
        {
            HeaderText = displayName,
            Name = columnName,
            Width = width,
            DataPropertyName = columnName,
            DisplayIndex = displayIndex,
            Visible = visible,
            Tag = tag,
            SortMode = DataGridViewColumnSortMode.Programmatic,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        };
        SetGridDefaultCellStyle(obj, dataType, decimalPlace, alignment);
        return obj;
    }

    private static void SetGridDefaultCellStyle(DataGridViewColumn dtColumn, string dataType, int decimalPlace = 0, int alignment = 16)
    {
        if (dtColumn != null)
        {
            dtColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            string text = dataType.ToLower();
            dtColumn.HeaderCell.Style.Alignment = (DataGridViewContentAlignment)alignment;
            dtColumn.DefaultCellStyle.Alignment = (DataGridViewContentAlignment)alignment;
            switch (text)
            {
                case "int64":
                case "int32":
                case "int16":
                    dtColumn.ValueType = typeof(int);
                    break;
                case "double":
                case "float":
                    dtColumn.ValueType = typeof(double);
                    ApplyNumberFormatting(dtColumn, decimalPlace);
                    break;
                case "datetime":
                    dtColumn.ValueType = typeof(DateTime);
                    dtColumn.DefaultCellStyle.Format = "dd/MM/yyyy";
                    break;
                default:
                    dtColumn.ValueType = typeof(string);
                    break;
            }
        }
    }

    private static void ApplyNumberFormatting(DataGridViewColumn dtColumn, int decimalPlace)
    {
        string text = "N" + ((decimalPlace >= 0) ? decimalPlace.ToString() : 2);
        dtColumn.DefaultCellStyle.Format = text;
    }

    /// <summary>
    /// A HashSet containing all the non-nullable integer types in .NET.
    /// </summary>
    private static readonly HashSet<Type> numericTypes = new()
    {
        typeof(sbyte), typeof(byte), typeof(short), typeof(ushort),
        typeof(int), typeof(uint), typeof(long), typeof(ulong),
        typeof(Int16), typeof(UInt16), typeof(Int32), typeof(UInt32),
        typeof(Int64), typeof(UInt64),
        typeof(double), typeof(float), typeof(decimal)
    };

    /// <summary>
    /// Determines whether the specified <see cref="DataGridViewColumn"/> contains numeric data.
    /// This method checks the column's <see cref="DataGridViewColumn.ValueType"/> to see if it
    /// is one of the common numeric types (byte, sbyte, short, ushort, int, uint, long, ulong, float, double, decimal).
    /// </summary>
    /// <param name="column">The <see cref="DataGridViewColumn"/> to check.</param>
    /// <returns>
    /// <c>true</c> if the column's <see cref="DataGridViewColumn.ValueType"/> is a numeric type;
    /// otherwise, <c>false</c>.
    /// </returns>
    public static bool IsNumericColumn(this DataGridViewColumn? column)
    {
        if (column == null || column.ValueType == null) return false;
        Type nonNullableType = Nullable.GetUnderlyingType(column.ValueType) ?? column.ValueType;
        return numericTypes.Contains(nonNullableType);
    }

    /// <summary>
    /// Checks if a given <see cref="Type"/> represents an numeric type. This method
    /// considers both nullable (e.g., <c>int?</c>) and non-nullable (e.g., <c>int</c>)
    /// integer types.
    /// </summary>
    /// <param name="type">The <see cref="Type"/> to check.</param>
    /// <returns><c>true</c> if the <paramref name="type"/> is an numeric type (or a nullable numeric type); otherwise, <c>false</c>.</returns>
    public static bool IsNumeric(this Type? type) =>
        type is not null && numericTypes.Contains(Nullable.GetUnderlyingType(type) ?? type);

    /// <summary>
    /// A HashSet containing all the non-nullable integer types in .NET.
    /// </summary>
    private static readonly HashSet<Type> IntegerTypes = new()
    {
        typeof(sbyte), typeof(byte), typeof(short), typeof(ushort),
        typeof(int), typeof(uint), typeof(long), typeof(ulong),
        typeof(Int16), typeof(UInt16), typeof(Int32), typeof(UInt32),
        typeof(Int64), typeof(UInt64)
    };

    /// <summary>
    /// A HashSet containing all the non-nullable floating-point number types in .NET.
    /// </summary>
    private static readonly HashSet<Type> FloatingPointTypes = new()
    {
        typeof(double), typeof(float), typeof(decimal)
    };

    /// <summary>
    /// Checks if a given <see cref="Type"/> represents an integer type. This method
    /// considers both nullable (e.g., <c>int?</c>) and non-nullable (e.g., <c>int</c>)
    /// integer types.
    /// </summary>
    /// <param name="type">The <see cref="Type"/> to check.</param>
    /// <returns><c>true</c> if the <paramref name="type"/> is an integer type (or a nullable integer type); otherwise, <c>false</c>.</returns>
    public static bool IsInteger(this Type? type) =>
        type is not null && IntegerTypes.Contains(Nullable.GetUnderlyingType(type) ?? type);

    /// <summary>
    /// Checks if a given <see cref="Type"/> represents a floating-point number type.
    /// This method considers both nullable (e.g., <c>double?</c>) and non-nullable
    /// (e.g., <c>double</c>, <c>float</c>, <c>decimal</c>) floating-point types.
    /// </summary>
    /// <param name="type">The <see cref="Type"/> to check.</param>
    /// <returns><c>true</c> if the <paramref name="type"/> is a floating-point number type (or a nullable floating-point type); otherwise, <c>false</c>.</returns>
    public static bool IsDouble(this Type? type) =>
        type is not null && FloatingPointTypes.Contains(Nullable.GetUnderlyingType(type) ?? type);

    /// <summary>
    /// Converts an object to an integer. If conversion fails, returns 0.
    /// </summary>
    /// <param name="obj">The object to convert.</param>
    /// <returns>An integer representation of the object, or 0 if conversion fails.</returns>
    public static int ToInteger(this object? obj)
    {
        if (obj == null) return 0;
        if (int.TryParse(obj.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out int result))
        {
            return result;
        }
        return 0;
    }

}