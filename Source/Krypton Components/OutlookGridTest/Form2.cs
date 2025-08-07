using System.Data;

using Krypton.Toolkit;

namespace OutlookGridTest
{
    public partial class Form2 : KryptonForm
    {

        #region Variables

        private readonly Random _random = new(); // Ensure _random is initialized
        private int _numberOfProductsToGenerate = 10; // Define common data generation parameters, Easily change for all methods

        private readonly string[] _categories = { "Electronics", "Furniture", "Books", "Apparel", "Home Goods", "Office Supplies", "Outdoor & Garden" };
        private readonly string[] _namePrefixes = { "Super", "Mega", "Ultra", "Smart", "Eco", "Pro", "Elite", "Compact" };
        private readonly string[] _nameSuffixes = { "X", "Plus", "Max", "Series", "Edition", "Go", "Connect", "Master" };
        private readonly string[] _commonProductWords = { "Laptop", "Mouse", "Keyboard", "Chair", "Monitor", "Desk", "Book", "Shirt", "Lamp", "Pen", "Tablet", "Speaker", "Camera", "Headphones" };

        #endregion Variables

        #region Identity

        public Form2()
        {
            InitializeComponent();

            kryptonAllInOneGrid1.OutlookGrid.OnGridColumnCreating += OutlookGrid_OnGridColumnCreating;
            kryptonAllInOneGrid1.OutlookGrid.OnInternalColumnCreating += OutlookGrid_OnInternalColumnCreating;

            kryptonOutlookGrid1.OnGridColumnCreating += OutlookGrid_OnGridColumnCreating;
            kryptonOutlookGrid1.OnInternalColumnCreating += OutlookGrid_OnInternalColumnCreating;

            outlookGrid1.OnGridColumnCreating += OutlookGrid_OnGridColumnCreating;
            outlookGrid1.OnInternalColumnCreating += OutlookGrid_OnInternalColumnCreating;
        }

        #endregion Identity

        #region Common Events

        private void Form2_Load(object sender, EventArgs e)
        {
            //DataTable dt = new DataTable("Tmp");
            //kryptonDataGridView1.Columns.Clear();
            //kryptonDataGridView1.AutoGenerateColumns = false;
            //KryptonDataGridViewRatingColumn Column12 = new KryptonDataGridViewRatingColumn()
            //{
            //    DataPropertyName = "Images",
            //    HeaderText = "Images int",
            //    RatingMaximum = 10,
            //    SortMode = DataGridViewColumnSortMode.Automatic
            //};

            //kryptonDataGridView1.Columns.Add(Column12);

            //dt = new();
            //dt.Columns.Add("Images", typeof(byte));

            //dt.Rows.Add(((byte)1));
            //dt.Rows.Add(((byte)2));
            //dt.Rows.Add(((byte)3));
            //dt.Rows.Add(((byte)4));

            //kryptonDataGridView1.DataSource = dt;

            //return;
            List<ProductDto> dataSource = GenerateProductData(_numberOfProductsToGenerate);
            AfterSetDataSource(dataSource);
            SetButtonText();
        }

        private void Form2_Shown(object sender, EventArgs e)
        {
            outlookGrid1.FitColumnsToWidth();
            kryptonOutlookGrid1.FitColumnsToWidth();
            kryptonAllInOneGrid1.OutlookGrid.FitColumnsToWidth();
        }

        private void TxtRecords_TextChanged(object sender, EventArgs e)
        {
            var rec = TxtRecords.Text.ToInteger();
            if (rec <= 0)
                _numberOfProductsToGenerate = 10;
            else
                _numberOfProductsToGenerate = rec;
        }

        private void OutlookGrid_OnInternalColumnCreating(object? sender, EventArgs e)
        {
            OutlookGridColumn column = (OutlookGridColumn)sender!;
            if (column.Name == "StockQuantity")
            {
                if (column.DataGridViewColumn.DataGridView!.Name == outlookGrid1.Name)
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
                column.Width = 70;
                if (column.ValueType.IsDouble())
                {
                    column.DefaultCellStyle.Format = "N2";
                }
            }
            else if (column.ValueType == typeof(DateTime))
            {
                // Apply specific formatting for DateTime columns
                column.DefaultCellStyle.Format = "dd/MM/yyyy";
                column.Width = 100;
            }
            else if (column.ValueType == typeof(bool))
            {
                column.Width = 70;
            }
            else
            {
                column.Width = 125;
            }
        }

        #endregion Common Events

        #region Button Spec

        private void buttonSpecHeaderGroup2_Click(object sender, EventArgs e)
        {
            LoadDataTableData();
        }

        private void buttonSpecHeaderGroup1_Click(object sender, EventArgs e)
        {
            KryptonMessageBox.Show("Data export successfully to Path: XXX.");
        }

        private void buttonSpecHeaderGroup3_Click(object sender, EventArgs e)
        {
            KryptonMessageBox.Show("Configuration saved successfully.");
        }

        #endregion Button Spec

        #region Buttons

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
            outlookGrid1.ReadOnly = !outlookGrid1.ReadOnly;
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
            HideIdColumn();
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
            SetButtonText();
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

        private void HideIdColumn()
        {
            bool isHidden = BtnShowIdInColumnContext.Text == "Show Id In Col Context";
            var idCol = outlookGrid1.FindFromColumnName("Id")!;
            idCol?.AvailableInContextMenu = isHidden;

            idCol = kryptonOutlookGrid1.FindFromColumnName("Id")!;
            idCol?.AvailableInContextMenu = isHidden;

            idCol = kryptonAllInOneGrid1.OutlookGrid.FindFromColumnName("Id")!;
            idCol?.AvailableInContextMenu = isHidden;

            if (isHidden)
                BtnShowIdInColumnContext.Text = "Hide Id In Col Context";
            else
                BtnShowIdInColumnContext.Text = "Show Id In Col Context";
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

            kryptonAllInOneGrid1.OutlookGrid.FillMode = kryptonOutlookGrid1.FillMode = outlookGrid1.FillMode;
            kryptonAllInOneGrid1.OutlookGrid.ShowSubTotal = kryptonOutlookGrid1.ShowSubTotal = outlookGrid1.ShowSubTotal;
            kryptonAllInOneGrid1.OutlookGrid.ShowGrandTotal = kryptonOutlookGrid1.ShowGrandTotal = outlookGrid1.ShowGrandTotal;
            kryptonAllInOneGrid1.OutlookGrid.ShowColumnFilter = kryptonOutlookGrid1.ShowColumnFilter = outlookGrid1.ShowColumnFilter;
            kryptonAllInOneGrid1.OutlookGrid.EnableSearchOnKeyPress = kryptonOutlookGrid1.EnableSearchOnKeyPress = outlookGrid1.EnableSearchOnKeyPress;
            kryptonAllInOneGrid1.OutlookGrid.HighlightSearchText = kryptonOutlookGrid1.HighlightSearchText = outlookGrid1.HighlightSearchText;
            kryptonAllInOneGrid1.OutlookGrid.SelectionMode = kryptonOutlookGrid1.SelectionMode = outlookGrid1.SelectionMode;
            kryptonAllInOneGrid1.OutlookGrid.ReadOnly = kryptonOutlookGrid1.ReadOnly = outlookGrid1.ReadOnly;
        }

        #endregion

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

                byte rating = Convert.ToByte(Math.Round(_random.NextDouble() * 5.0, 1)); // Rating between 0.0 and 5.0, one decimal place
                if (rating < 1) rating = Convert.ToByte(1); // Ensure a minimum rating of 1.0 for more realistic data

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
            List<ProductDto> dataSource = GenerateProductData(_numberOfProductsToGenerate);

            outlookGrid1.SetDataSource(dataSource);
            kryptonOutlookGrid1.SetDataSource(dataSource);
            kryptonAllInOneGrid1.OutlookGrid.SetDataSource(dataSource);
            kryptonDataGridView1.DataSource = dataSource;
            AfterSetDataSource();
        }

        private void LoadDataTableData()
        {
            // 1. Create a new DataTable
            DataTable dataSource = new DataTable("Products");

            // 2. Define the columns for the DataTable, specifying their types
            dataSource.Columns.Add("Id", typeof(int));
            dataSource.Columns.Add("Name", typeof(string));
            dataSource.Columns.Add("Category", typeof(string));
            dataSource.Columns.Add("Price", typeof(decimal)); // Use decimal for currency
            dataSource.Columns.Add("StockQuantity", typeof(int));
            dataSource.Columns.Add("LastRestockDate", typeof(DateTime));
            dataSource.Columns.Add("IsAvailable", typeof(bool));
            dataSource.Columns.Add("Rating", typeof(byte));
            // productsTable.Columns.Add("ProductImage", typeof(Image)); // Add if you intend to use images

            // Get generated product data
            List<ProductDto> products = GenerateProductData(_numberOfProductsToGenerate);

            // 3. Add rows to the DataTable from the generated data
            foreach (var product in products)
            {
                dataSource.Rows.Add(
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

            outlookGrid1.SetDataSource(dataSource);
            //outlookGrid1.AutoResizeColumnsToFit();
            kryptonOutlookGrid1.SetDataSource(dataSource);
            kryptonAllInOneGrid1.OutlookGrid.SetDataSource(dataSource);
            kryptonDataGridView1.DataSource = dataSource;
            AfterSetDataSource();
        }

        private void LoadRawArrayData()
        {
            List<object[]> dataSource = new List<object[]>();

            // Get generated product data
            List<ProductDto> products = GenerateProductData(_numberOfProductsToGenerate);

            // Convert ProductDto objects to object arrays
            foreach (var product in products)
            {
                dataSource.Add(new object[]
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

            outlookGrid1.SetDataSource(dataSource);
            //outlookGrid1.AutoResizeColumnsToFit();
            kryptonOutlookGrid1.SetDataSource(dataSource);
            kryptonAllInOneGrid1.OutlookGrid.SetDataSource(dataSource);
            kryptonDataGridView1.DataSource = dataSource;
            AfterSetDataSource();
        }

        private void LoadDictionaryData()
        {
            List<Dictionary<string, object>> dataSource = new List<Dictionary<string, object>>();

            // Get generated product data
            List<ProductDto> products = GenerateProductData(_numberOfProductsToGenerate);

            // Convert ProductDto objects to Dictionaries
            foreach (var product in products)
            {
                dataSource.Add(new Dictionary<string, object>
                {
                    { "Id", product.Id },
                    { "Name", product.Name }, // Use consistent key names as desired by your grid
                    { "Category", product.Category },
                    { "Price", product.Price },
                    { "StockQuantity", product.StockQuantity },
                    { "LastRestockDate", product.LastRestockDate },
                    { "IsAvailable", product.IsAvailable },
                    { "Rating", product.Rating }
                    // { "ProductImage", product.ProductImage } // Include if you want to pass images to dictionary
                });
            }

            outlookGrid1.SetDataSource(dataSource);
            //outlookGrid1.AutoResizeColumnsToFit();
            kryptonOutlookGrid1.SetDataSource(dataSource);
            kryptonAllInOneGrid1.OutlookGrid.SetDataSource(dataSource);
            kryptonDataGridView1.DataSource = dataSource;
            AfterSetDataSource();
        }

        private void AfterSetDataSource(object? dataSource = null)
        {
            if (dataSource != null)
            {
                outlookGrid1.SetDataSource(dataSource);
                //outlookGrid1.AutoResizeColumnsToFit();
                kryptonOutlookGrid1.SetDataSource(dataSource);
                kryptonAllInOneGrid1.OutlookGrid.SetDataSource(dataSource);
                kryptonDataGridView1.DataSource = dataSource;
            }

            // Apply formatting for the Price column
            if (outlookGrid1.Columns.Contains("Price"))
            {
                outlookGrid1.Columns["Price"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns["Price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                outlookGrid1.Columns["Price"].DefaultCellStyle.Format = "N2";
            }
            if (kryptonOutlookGrid1.Columns.Contains("Price"))
            {
                kryptonOutlookGrid1.Columns["Price"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                kryptonOutlookGrid1.Columns["Price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                kryptonOutlookGrid1.Columns["Price"].DefaultCellStyle.Format = "N2";
            }
            if (kryptonAllInOneGrid1.OutlookGrid.Columns.Contains("Price"))
            {
                kryptonAllInOneGrid1.OutlookGrid.Columns["Price"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                kryptonAllInOneGrid1.OutlookGrid.Columns["Price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                kryptonAllInOneGrid1.OutlookGrid.Columns["Price"].DefaultCellStyle.Format = "N2";
            }

            if (outlookGrid1.Columns.Contains("Id"))
            {
                outlookGrid1.Columns["Id"].Visible = false;
                kryptonOutlookGrid1.Columns["Id"].Visible = false;
                kryptonAllInOneGrid1.Columns["Id"].Visible = false;
                BtnShowIdInColumnContext.PerformClick();
            }

            outlookGrid1.FitColumnsToWidth();
            kryptonOutlookGrid1.FitColumnsToWidth();
            //kryptonAllInOneGrid1.OutlookGrid.Refresh();
            kryptonAllInOneGrid1.OutlookGrid.FitColumnsToWidth();
        }

        #endregion Get Data Methods

    }
}
