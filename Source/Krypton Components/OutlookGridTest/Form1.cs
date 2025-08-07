using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Xml.Linq;

using Krypton.Toolkit;

namespace OutlookGridTest
{
    public partial class Form1 : Form
    {

        private readonly Random _random = new(); // Ensure _random is initialized

        // Define common data generation parameters
        private const int NumberOfProductsToGenerate = 10000;//10000; // Easily change for all methods

        private readonly string[] _categories = { "Electronics", "Furniture", "Books", "Apparel", "Home Goods", "Office Supplies", "Outdoor & Garden" };
        private readonly string[] _namePrefixes = { "Super", "Mega", "Ultra", "Smart", "Eco", "Pro", "Elite", "Compact" };
        private readonly string[] _nameSuffixes = { "X", "Plus", "Max", "Series", "Edition", "Go", "Connect", "Master" };
        private readonly string[] _commonProductWords = { "Laptop", "Mouse", "Keyboard", "Chair", "Monitor", "Desk", "Book", "Shirt", "Lamp", "Pen", "Tablet", "Speaker", "Camera", "Headphones" };

        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;
            kryptonExtraGrid1.OutlookGrid.OnGridColumnCreating += OutlookGrid_OnGridColumnCreating;
            kryptonExtraGrid1.OutlookGrid.OnInternalColumnCreating += OutlookGrid_OnInternalColumnCreating;
        }

        private void OutlookGrid_OnInternalColumnCreating(object? sender, EventArgs e)
        {
            OutlookGridColumn column = (OutlookGridColumn)sender!;
            if (column.Name == "StockQuantity")
            {
                column.AggregationFormatString = "{1} of {0}: {2}";
                column.AggregationType = KryptonOutlookGridAggregationType.Sum;
            }
            else if (column.Name == "Category")
            {
                column.GroupIndex = 0;
                column.SortIndex = 0;
                column.SortDirection = SortOrder.Ascending;
            }
        }

        private void OutlookGrid_OnGridColumnCreating(object? sender, EventArgs e)
        {
            DataGridViewColumn column = (DataGridViewColumn)sender!;
            if (column.Name == "Rating")
            {
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            else if (column.ValueType.IsNumeric())
            {
                // Apply specific formatting for numeric columns
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                column.Width = 100;
                if (column.ValueType.IsDouble())
                {
                    column.DefaultCellStyle.Format = "N2";
                }
            }
            else if (column.ValueType == typeof(DateTime))
            {
                // Apply specific formatting for DateTime columns
                column.DefaultCellStyle.Format = "dd/MM/yyyy";
                column.Width = 125;
            }
            else
            {
                column.Width = 250;
            }
        }

        private void Form1_Load(object? sender, EventArgs e)
        {
            outlookGrid1.FillMode = GridFillMode.GroupsAndNodes;
            //FillWithCurrentMethod();
            SetButtonText();
            SetTestData();
        }

        private BindingSource bsData = new();
        private DataTable _dataTable = null!;
        private DataSet _dataSet = null!;
        private void SetTestData()
        {
            _dataTable = new DataTable();
            _dataSet = new DataSet();
            bsData.DataSource = _dataSet;
            _dataTable = _dataSet.Tables.Add("TableTest");
            _dataTable.Columns.Add("int", typeof(int));
            _dataTable.Columns.Add("decimal", typeof(decimal));
            _dataTable.Columns.Add("double", typeof(double));
            _dataTable.Columns.Add("date", typeof(DateTime));
            _dataTable.Columns.Add("datetime", typeof(DateTime));
            _dataTable.Columns.Add("string", typeof(string));
            _dataTable.Columns.Add("boolean", typeof(bool));
            _dataTable.Columns.Add("guid", typeof(Guid));
            //_dataTable.Columns.Add("image", typeof(Bitmap));
            _dataTable.Columns.Add("timespan", typeof(TimeSpan));

            bsData.DataSource = _dataSet;
            bsData.DataMember = _dataTable.TableName;

            //kryptonExtraGrid1.OutlookGrid.DataSource = bsData;
            AddTestData();
            kryptonExtraGrid1.OutlookGrid.SetDataSource(bsData);
        }

        private void AddTestData()
        {
            Random r = new Random();
            /*Image[] sampleimages = new Image[2];
            sampleimages[0] = Image.FromFile(Path.Combine(Application.StartupPath, "flag-green_24.png"));
            sampleimages[1] = Image.FromFile(Path.Combine(Application.StartupPath, "flag-red_24.png"));*/

            int maxMinutes = (int)((TimeSpan.FromHours(20) - TimeSpan.FromHours(10)).TotalMinutes);
            for (int i = 0; i < 1000; i++)
            {
                object[] newrow = new object[] {
                    i,
                    Math.Round((decimal)i*2/3, 6),
                    Math.Round(i % 2 == 0 ? (double)i*2/3 : (double)i/2, 6),
                    DateTime.Today.AddHours(i*2).AddHours(i%2 == 0 ?i*10+1:0).AddMinutes(i%2 == 0 ?i*10+1:0).AddSeconds(i%2 == 0 ?i*10+1:0).AddMilliseconds(i%2 == 0 ?i*10+1:0).Date,
                    DateTime.Today.AddHours(i*2).AddHours(i%2 == 0 ?i*10+1:0).AddMinutes(i%2 == 0 ?i*10+1:0).AddSeconds(i%2 == 0 ?i*10+1:0).AddMilliseconds(i%2 == 0 ?i*10+1:0),
                    i*2 % 3 == 0 ? null! : $"{i} str",
                    i % 2 == 0 ? true:false,
                    Guid.NewGuid(),
                    //null,
                    //sampleimages[r.Next(0, 2)],
                    TimeSpan.FromHours(10).Add(TimeSpan.FromMinutes(r.Next(maxMinutes)))
                };

                _dataTable.Rows.Add(newrow);
            }
        }

        private void FillWithCurrentMethod()
        {
            OutlookGridRow row = new();
            List<OutlookGridRow> l = [];
            outlookGrid1.ClearEverything();
            outlookGrid1.HideColumnOnGrouping = false;
            outlookGrid1.FillMode = GridFillMode.GroupsAndNodes;
            outlookGrid1.ShowLines = true;
            outlookGrid1.SuspendLayout();

            int dispIndex = 0;
            outlookGrid1.CreateOutlookGridColumn("Id", "Id", 100, dispIndex++, false);
            outlookGrid1.CreateOutlookGridColumn("Name", "Name", 100, dispIndex++, true);
            outlookGrid1.CreateOutlookGridColumn("Cat", "Cat", 125, dispIndex++, true, false, SortOrder.Ascending, 0);
            outlookGrid1.CreateOutlookGridColumn("Price", "Price", 125, dispIndex++, true, false, SortOrder.None, -1, DataGridViewContentAlignment.MiddleRight, 2, KryptonOutlookGridAggregationType.Sum, "double");
            outlookGrid1.CreateOutlookGridColumn("Quantity", "Quantity", 125, dispIndex++, true, false, SortOrder.None, -1, DataGridViewContentAlignment.MiddleRight, 2, KryptonOutlookGridAggregationType.Sum, "double");
            outlookGrid1.CreateOutlookGridColumn("OrdDate", "OrdDate", 125, dispIndex++, true, false, SortOrder.None, -1, DataGridViewContentAlignment.MiddleLeft, -1, KryptonOutlookGridAggregationType.None, "datetime");
            /*outlookGrid1.Columns["Price"].ValueType = typeof(double);
            outlookGrid1.Columns["Quantity"].ValueType = typeof(double);
            outlookGrid1.Columns["OrdDate"].ValueType = typeof(DateTime);*/

            row = new OutlookGridRow();
            row.CreateCells(outlookGrid1, new object[] { 1, "Laptop", "Electronics", 1200.50m, 5, new DateTime(2025, 5, 10) });
            l.Add(row);
            row = new OutlookGridRow();
            row.CreateCells(outlookGrid1, new object[] { 2, "Mouse", "Electronics", 25.99m, 20, new DateTime(2025, 5, 15) });
            l.Add(row);
            row = new OutlookGridRow();
            row.CreateCells(outlookGrid1, new object[] { 3, "Keyboard", "Electronics", 75.00m, 12, new DateTime(2025, 5, 20) });
            l.Add(row);
            row = new OutlookGridRow();
            row.CreateCells(outlookGrid1, new object[] { 4, "Desk Chair", "Furniture", 150.00m, 8, new DateTime(2025, 5, 12) });
            l.Add(row);
            row = new OutlookGridRow();
            row.CreateCells(outlookGrid1, new object[] { 6, "Bookcase", "Furniture", 90.00m, 10, new DateTime(2025, 6, 5) });
            l.Add(row);
            row = new OutlookGridRow();
            row.CreateCells(outlookGrid1, new object[] { 7, "Monitor", "Electronics", 90.00m, 10, new DateTime(2025, 6, 5) });
            l.Add(row);

            var intCol = outlookGrid1.FindFromColumnName("Id")!;
            intCol.AvailableInContextMenu = false;

            outlookGrid1.ResumeLayout();
            outlookGrid1.AssignRows(l);
            outlookGrid1.ForceRefreshGroupBox();

            outlookGrid1.Fill();
        }

        private void BtnShowSubTotal_Click(object sender, EventArgs e)
        {
            outlookGrid1.ShowSubTotal = !outlookGrid1.ShowSubTotal;
            SetButtonText();
        }

        private void BtnShowGrandTotal_Click(object sender, EventArgs e)
        {
            outlookGrid1.ShowGrandTotal = !outlookGrid1.ShowGrandTotal;
            SetButtonText();
        }

        private void BtnShowColumnFilter_Click(object sender, EventArgs e)
        {
            outlookGrid1.ShowColumnFilter = !outlookGrid1.ShowColumnFilter;
            SetButtonText();
        }

        private void BtnEnableSearch_Click(object sender, EventArgs e)
        {
            outlookGrid1.EnableSearchOnKeyPress = !outlookGrid1.EnableSearchOnKeyPress;
            SetButtonText();
        }

        private void BtnHighlightSearch_Click(object sender, EventArgs e)
        {
            outlookGrid1.HighlightSearchText = !outlookGrid1.HighlightSearchText;
            outlookGrid1.Invalidate();
            SetButtonText();
        }

        private void BtnShowIdInColumnContext_Click(object sender, EventArgs e)
        {
            // reset for testing because it's work after context menu reset
            FillWithCurrentMethod();
            var intCol = outlookGrid1.FindFromColumnName("Id")!;
            if (BtnShowIdInColumnContext.Text == "Hide Id In Col Context")
                intCol.AvailableInContextMenu = false;
            else
                intCol.AvailableInContextMenu = true;
            SetButtonText();
        }

        private void BtnSelectionMode_Click(object sender, EventArgs e)
        {
            if (outlookGrid1.SelectionMode == DataGridViewSelectionMode.FullRowSelect)
            {
                outlookGrid1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            }
            else
            {
                outlookGrid1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            }
        }

        private void SetButtonText()
        {
            if (outlookGrid1.ShowSubTotal)
                BtnShowSubTotal.Text = "Hide Sub Total";
            else
                BtnShowSubTotal.Text = "Show Sub Total";

            if (outlookGrid1.ShowGrandTotal)
                BtnShowGrandTotal.Text = "Hide Grand Total";
            else
                BtnShowGrandTotal.Text = "Show Grand Total";

            if (outlookGrid1.ShowColumnFilter)
                BtnShowColumnFilter.Text = "Hide Column Filter";
            else
                BtnShowColumnFilter.Text = "Show Column Filter";

            if (outlookGrid1.EnableSearchOnKeyPress)
                BtnEnableSearch.Text = "Disable Search";
            else
                BtnEnableSearch.Text = "Enable Search";

            BtnHighlightSearch.Text = $"Highlight Search : {outlookGrid1.HighlightSearchText}";

            if (outlookGrid1.SelectionMode == DataGridViewSelectionMode.FullRowSelect)
            {
                BtnSelectionMode.Text = "Cell Select";
            }
            else
            {
                BtnSelectionMode.Text = "Full Row Select";
            }

            kryptonExtraGrid1.OutlookGrid.FillMode = outlookGrid1.FillMode;
            kryptonExtraGrid1.OutlookGrid.ShowSubTotal = outlookGrid1.ShowSubTotal;
            kryptonExtraGrid1.OutlookGrid.ShowGrandTotal = outlookGrid1.ShowGrandTotal;
            kryptonExtraGrid1.OutlookGrid.ShowColumnFilter = outlookGrid1.ShowColumnFilter;
            kryptonExtraGrid1.OutlookGrid.EnableSearchOnKeyPress = outlookGrid1.EnableSearchOnKeyPress;
            kryptonExtraGrid1.OutlookGrid.HighlightSearchText = outlookGrid1.HighlightSearchText;
            kryptonExtraGrid1.OutlookGrid.SelectionMode = outlookGrid1.SelectionMode;
            kryptonExtraGrid1.OutlookGrid.ReadOnly = outlookGrid1.ReadOnly;


            /*var intCol = outlookGrid1.FindFromColumnName("Id")!;
            if (intCol.AvailableInContextMenu)
                BtnShowIdInColumnContext.Text = "Hide Id In Col Context";
            else
                BtnShowIdInColumnContext.Text = "Show Id In Col Context";*/

        }


        private void BtnLoadDataTable_Click(object sender, EventArgs e)
        {
            LoadDataTableData();
        }

        private void BtnLoadListOfT_Click(object sender, EventArgs e)
        {
            LoadProductData();
        }

        private void BtnLoadListOfRawArray_Click(object sender, EventArgs e)
        {
            LoadRawArrayData();
        }

        private void BtnLoadDictionary_Click(object sender, EventArgs e)
        {
            LoadDictionaryData();
        }

        #region Get Data Methods

        /// <summary>
        /// Generates a list of synthetic product data.
        /// </summary>
        /// <param name="count">The number of products to generate.</param>
        /// <returns>A list of <see cref="ProductDto"/> objects.</returns>
        private List<ProductDto> GenerateProductData(int count)
        {
            List<ProductDto> products = new List<ProductDto>();

            for (int i = 1; i <= count; i++)
            {
                string category = _categories[_random.Next(_categories.Length)];
                string productName = $"{_namePrefixes[_random.Next(_namePrefixes.Length)]} {_commonProductWords[_random.Next(_commonProductWords.Length)]} {_nameSuffixes[_random.Next(_nameSuffixes.Length)]} {i}";

                decimal price = Math.Round((decimal)(_random.NextDouble() * 500) + 10, 2); // Price between 10.00 and 510.00
                int stockQuantity = _random.Next(1, 200); // Stock between 1 and 199

                // Dates within the last year or so, including current date
                DateTime lastUpdated = DateTime.Now.AddDays(-_random.Next(0, 365));
                bool isAvailable = _random.Next(0, 100) < 95; // 95% chance of being available

                double rating = Math.Round(_random.NextDouble() * 5.0, 1); // Rating between 0.0 and 5.0, one decimal place
                if (rating < 1.0) rating = 1.0; // Ensure a minimum rating of 1.0 for more realistic data

                Image? productImage = null; // Keep null for large data sets unless actual image data is needed for testing

                products.Add(new ProductDto(
                    i,
                    productName,
                    category,
                    price,
                    stockQuantity,
                    lastUpdated,
                    isAvailable,
                    rating,
                    productImage
                ));
            }
            return products;
        }

        private void LoadProductData()
        {
            List<ProductDto> products = GenerateProductData(NumberOfProductsToGenerate);

            // Example of setting data for your grid
            outlookGrid1.SetDataSource(products);
            outlookGrid1.FitColumnsToWidth();

            kryptonExtraGrid1.OutlookGrid.SetDataSource(products);
            kryptonExtraGrid1.OutlookGrid.FitColumnsToWidth();

            // Apply formatting after setting data source
            if (outlookGrid1.Columns.Contains("Price"))
            {
                outlookGrid1.Columns["Price"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns["Price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns["Price"].DefaultCellStyle.Format = "N2";
            }
        }

        private void LoadDataTableData()
        {
            // 1. Create a new DataTable
            DataTable productsTable = new DataTable("Products");

            // 2. Define the columns for the DataTable, specifying their types
            productsTable.Columns.Add("Id", typeof(int));
            productsTable.Columns.Add("Name", typeof(string));
            productsTable.Columns.Add("Category", typeof(string));
            productsTable.Columns.Add("Price", typeof(decimal)); // Use decimal for currency
            productsTable.Columns.Add("StockQuantity", typeof(int));
            productsTable.Columns.Add("LastRestockDate", typeof(DateTime));
            productsTable.Columns.Add("IsAvailable", typeof(bool));
            productsTable.Columns.Add("Rating", typeof(double));
            // productsTable.Columns.Add("ProductImage", typeof(Image)); // Add if you intend to use images

            // Get generated product data
            List<ProductDto> products = GenerateProductData(NumberOfProductsToGenerate);

            // 3. Add rows to the DataTable from the generated data
            foreach (var product in products)
            {
                productsTable.Rows.Add(
                    product.Id,
                    product.Name,
                    product.Category,
                    product.Price,
                    product.StockQuantity,
                    product.LastRestockDate,
                    product.IsAvailable,
                    product.Rating
                // Add product.ProductImage if you have an Image column
                );
            }

            outlookGrid1.DataSource = productsTable;
            outlookGrid1.FitColumnsToWidth();

            kryptonExtraGrid1.OutlookGrid.SetDataSource(productsTable);
            kryptonExtraGrid1.OutlookGrid.FitColumnsToWidth();

            // Apply formatting for the Price column
            if (outlookGrid1.Columns.Contains("Price"))
            {
                outlookGrid1.Columns["Price"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns["Price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns["Price"].DefaultCellStyle.Format = "N2";
            }
        }

        private void LoadRawArrayData()
        {
            List<object[]> rawData = new List<object[]>();

            // Get generated product data
            List<ProductDto> products = GenerateProductData(NumberOfProductsToGenerate);

            // Convert ProductDto objects to object arrays
            foreach (var product in products)
            {
                rawData.Add(new object[]
                {
                product.Id,
                product.Name,
                product.Category,
                product.Price,
                product.StockQuantity,
                product.LastRestockDate,
                product.IsAvailable,
                product.Rating,
                    // product.ProductImage // Include if you want to pass images to raw array
                });
            }

            outlookGrid1.SetDataSource(rawData);
            outlookGrid1.FitColumnsToWidth();

            kryptonExtraGrid1.OutlookGrid.SetDataSource(rawData);
            kryptonExtraGrid1.OutlookGrid.FitColumnsToWidth();
            // Apply formatting for the Price column (assuming column index or name)
            // If your grid auto-generates columns from raw array, column names might not be available directly.
            // You might need to rely on column index or inspect the generated columns.
            // Assuming "Price" will be at index 3 based on your array structure:
            if (outlookGrid1.Columns.Count > 3) // Check if column exists
            {
                outlookGrid1.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns[3].DefaultCellStyle.Format = "N2";
            }
            Console.WriteLine($"Generated and loaded {rawData.Count} products for raw object array.");
        }

        private void LoadDictionaryData()
        {
            List<Dictionary<string, object>> dictionaryData = new List<Dictionary<string, object>>();

            // Get generated product data
            List<ProductDto> products = GenerateProductData(NumberOfProductsToGenerate);

            // Convert ProductDto objects to Dictionaries
            foreach (var product in products)
            {
                dictionaryData.Add(new Dictionary<string, object>
            {
                { "Id", product.Id },
                { "ProductName", product.Name }, // Use consistent key names as desired by your grid
                { "ProductCategory", product.Category },
                { "UnitPrice", product.Price },
                { "QtyInStock", product.StockQuantity },
                { "RestockDate", product.LastRestockDate },
                { "Available", product.IsAvailable },
                { "CustomerRating", product.Rating }
                // { "ProductImage", product.ProductImage } // Include if you want to pass images to dictionary
            });
            }

            outlookGrid1.SetDataSource(dictionaryData);
            outlookGrid1.FitColumnsToWidth();

            kryptonExtraGrid1.OutlookGrid.SetDataSource(dictionaryData);
            kryptonExtraGrid1.OutlookGrid.FitColumnsToWidth();

            // Apply formatting for the Price column
            if (outlookGrid1.Columns.Contains("UnitPrice")) // Use the key name from your dictionary
            {
                outlookGrid1.Columns["UnitPrice"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns["UnitPrice"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns["UnitPrice"].DefaultCellStyle.Format = "N2";
            }
        }

        #endregion Get Data Methods

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

    public class ProductDto
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
        public decimal Price { get; set; } // Using decimal for currency
        public int StockQuantity { get; set; }
        public DateTime LastRestockDate { get; set; }
        public bool IsAvailable { get; set; }
        public double Rating { get; set; } // For a double type
        public Image? Thumbnail { get; set; } // For an Image type (needs System.Drawing)

        // Constructor for easy initialization
        public ProductDto(int id, string name, string category, decimal price, int stockQuantity, DateTime lastRestockDate, bool isAvailable, double rating, Image? thumbnail = null)
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
        }
    }

}
