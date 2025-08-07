namespace OutlookGridTest
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.BtnEnableSearch = new Krypton.Toolkit.KryptonButton();
            this.BtnShowColumnFilter = new Krypton.Toolkit.KryptonButton();
            this.BtnShowGrandTotal = new Krypton.Toolkit.KryptonButton();
            this.BtnShowSubTotal = new Krypton.Toolkit.KryptonButton();
            this.outlookGrid1 = new Krypton.Toolkit.KryptonOutlookGrid();
            this.kryptonContextMenu1 = new Krypton.Toolkit.KryptonContextMenu();
            this.kryptonContextMenuItem1 = new Krypton.Toolkit.KryptonContextMenuItem();
            this.kryptonContextMenuCheckBox1 = new Krypton.Toolkit.KryptonContextMenuCheckBox();
            this.kryptonContextMenuHeading1 = new Krypton.Toolkit.KryptonContextMenuHeading();
            this.kryptonContextMenuRadioButton1 = new Krypton.Toolkit.KryptonContextMenuRadioButton();
            this.kryptonContextMenuCheckButton1 = new Krypton.Toolkit.KryptonContextMenuCheckButton();
            this.kryptonContextMenuLinkLabel1 = new Krypton.Toolkit.KryptonContextMenuLinkLabel();
            this.kryptonContextMenuColorColumns1 = new Krypton.Toolkit.KryptonContextMenuColorColumns();
            this.kryptonContextMenuImageSelect1 = new Krypton.Toolkit.KryptonContextMenuImageSelect();
            this.kryptonContextMenuMonthCalendar2 = new Krypton.Toolkit.KryptonContextMenuMonthCalendar();
            this.kryptonContextMenuMonthCalendar1 = new Krypton.Toolkit.KryptonContextMenuMonthCalendar();
            this.kryptonContextMenuItem2 = new Krypton.Toolkit.KryptonContextMenuItem();
            this.BtnHighlightSearch = new Krypton.Toolkit.KryptonButton();
            this.BtnShowIdInColumnContext = new Krypton.Toolkit.KryptonButton();
            this.kryptonButton1 = new Krypton.Toolkit.KryptonButton();
            this.BtnLoadListOfT = new Krypton.Toolkit.KryptonButton();
            this.BtnLoadListOfRawArray = new Krypton.Toolkit.KryptonButton();
            this.BtnLoadDictionary = new Krypton.Toolkit.KryptonButton();
            this.BtnSelectionMode = new Krypton.Toolkit.KryptonButton();
            this.kryptonTableLayoutPanel1 = new Krypton.Toolkit.KryptonTableLayoutPanel();
            this.kryptonExtraGrid1 = new Krypton.Toolkit.KryptonAllInOneGrid();
            this.buttonSpecHeaderGroup1 = new Krypton.Toolkit.ButtonSpecHeaderGroup();
            this._textBox = new Krypton.Toolkit.KryptonTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.outlookGrid1)).BeginInit();
            this.kryptonTableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonExtraGrid1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonExtraGrid1.OutlookGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonExtraGrid1.Panel)).BeginInit();
            this.kryptonExtraGrid1.Panel.SuspendLayout();
            this.kryptonExtraGrid1.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnEnableSearch
            // 
            this.BtnEnableSearch.Location = new System.Drawing.Point(405, 12);
            this.BtnEnableSearch.Name = "BtnEnableSearch";
            this.BtnEnableSearch.Size = new System.Drawing.Size(125, 25);
            this.BtnEnableSearch.TabIndex = 4;
            this.BtnEnableSearch.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnEnableSearch.Values.Text = "Enable Search";
            this.BtnEnableSearch.Click += new System.EventHandler(this.BtnEnableSearch_Click);
            // 
            // BtnShowColumnFilter
            // 
            this.BtnShowColumnFilter.Location = new System.Drawing.Point(274, 12);
            this.BtnShowColumnFilter.Name = "BtnShowColumnFilter";
            this.BtnShowColumnFilter.Size = new System.Drawing.Size(125, 25);
            this.BtnShowColumnFilter.TabIndex = 3;
            this.BtnShowColumnFilter.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnShowColumnFilter.Values.Text = "Show Column Filter";
            this.BtnShowColumnFilter.Click += new System.EventHandler(this.BtnShowColumnFilter_Click);
            // 
            // BtnShowGrandTotal
            // 
            this.BtnShowGrandTotal.Location = new System.Drawing.Point(143, 12);
            this.BtnShowGrandTotal.Name = "BtnShowGrandTotal";
            this.BtnShowGrandTotal.Size = new System.Drawing.Size(125, 25);
            this.BtnShowGrandTotal.TabIndex = 2;
            this.BtnShowGrandTotal.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnShowGrandTotal.Values.Text = "Show Grand Total";
            this.BtnShowGrandTotal.Click += new System.EventHandler(this.BtnShowGrandTotal_Click);
            // 
            // BtnShowSubTotal
            // 
            this.BtnShowSubTotal.Location = new System.Drawing.Point(12, 12);
            this.BtnShowSubTotal.Name = "BtnShowSubTotal";
            this.BtnShowSubTotal.Size = new System.Drawing.Size(125, 25);
            this.BtnShowSubTotal.TabIndex = 1;
            this.BtnShowSubTotal.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnShowSubTotal.Values.Text = "Show Sub Total";
            this.BtnShowSubTotal.Click += new System.EventHandler(this.BtnShowSubTotal_Click);
            // 
            // outlookGrid1
            // 
            this.outlookGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.outlookGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.outlookGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.outlookGrid1.EnableSearchOnKeyPress = true;
            this.outlookGrid1.FillMode = Krypton.Toolkit.GridFillMode.GroupsOnly;
            this.outlookGrid1.Location = new System.Drawing.Point(3, 3);
            this.outlookGrid1.Name = "outlookGrid1";
            this.outlookGrid1.ReadOnly = true;
            this.outlookGrid1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.outlookGrid1.ShowColumnFilter = true;
            this.outlookGrid1.ShowGrandTotal = true;
            this.outlookGrid1.ShowLines = false;
            this.outlookGrid1.ShowSubTotal = true;
            this.outlookGrid1.Size = new System.Drawing.Size(999, 270);
            this.outlookGrid1.TabIndex = 0;
            // 
            // kryptonContextMenu1
            // 
            this.kryptonContextMenu1.Items.AddRange(new Krypton.Toolkit.KryptonContextMenuItemBase[] {
            this.kryptonContextMenuItem1,
            this.kryptonContextMenuCheckBox1,
            this.kryptonContextMenuHeading1,
            this.kryptonContextMenuRadioButton1,
            this.kryptonContextMenuCheckButton1,
            this.kryptonContextMenuLinkLabel1,
            this.kryptonContextMenuColorColumns1,
            this.kryptonContextMenuImageSelect1,
            this.kryptonContextMenuMonthCalendar2});
            // 
            // kryptonContextMenuColorColumns1
            // 
            this.kryptonContextMenuColorColumns1.SelectedColor = System.Drawing.Color.Empty;
            // 
            // BtnHighlightSearch
            // 
            this.BtnHighlightSearch.Location = new System.Drawing.Point(536, 12);
            this.BtnHighlightSearch.Name = "BtnHighlightSearch";
            this.BtnHighlightSearch.Size = new System.Drawing.Size(150, 25);
            this.BtnHighlightSearch.TabIndex = 5;
            this.BtnHighlightSearch.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnHighlightSearch.Values.Text = "Highlight Search: true";
            this.BtnHighlightSearch.Click += new System.EventHandler(this.BtnHighlightSearch_Click);
            // 
            // BtnShowIdInColumnContext
            // 
            this.BtnShowIdInColumnContext.Location = new System.Drawing.Point(692, 12);
            this.BtnShowIdInColumnContext.Name = "BtnShowIdInColumnContext";
            this.BtnShowIdInColumnContext.Size = new System.Drawing.Size(150, 25);
            this.BtnShowIdInColumnContext.TabIndex = 6;
            this.BtnShowIdInColumnContext.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnShowIdInColumnContext.Values.Text = "Show Id In Col Context";
            this.BtnShowIdInColumnContext.Click += new System.EventHandler(this.BtnShowIdInColumnContext_Click);
            // 
            // kryptonButton1
            // 
            this.kryptonButton1.Location = new System.Drawing.Point(12, 43);
            this.kryptonButton1.Name = "kryptonButton1";
            this.kryptonButton1.Size = new System.Drawing.Size(125, 25);
            this.kryptonButton1.TabIndex = 7;
            this.kryptonButton1.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.kryptonButton1.Values.Text = "BtnLoadDataTable";
            this.kryptonButton1.Click += new System.EventHandler(this.BtnLoadDataTable_Click);
            // 
            // BtnLoadListOfT
            // 
            this.BtnLoadListOfT.Location = new System.Drawing.Point(143, 43);
            this.BtnLoadListOfT.Name = "BtnLoadListOfT";
            this.BtnLoadListOfT.Size = new System.Drawing.Size(125, 25);
            this.BtnLoadListOfT.TabIndex = 8;
            this.BtnLoadListOfT.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnLoadListOfT.Values.Text = "Load List<T>";
            this.BtnLoadListOfT.Click += new System.EventHandler(this.BtnLoadListOfT_Click);
            // 
            // BtnLoadListOfRawArray
            // 
            this.BtnLoadListOfRawArray.Location = new System.Drawing.Point(274, 43);
            this.BtnLoadListOfRawArray.Name = "BtnLoadListOfRawArray";
            this.BtnLoadListOfRawArray.Size = new System.Drawing.Size(125, 25);
            this.BtnLoadListOfRawArray.TabIndex = 9;
            this.BtnLoadListOfRawArray.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnLoadListOfRawArray.Values.Text = "Load List<object[]>";
            this.BtnLoadListOfRawArray.Click += new System.EventHandler(this.BtnLoadListOfRawArray_Click);
            // 
            // BtnLoadDictionary
            // 
            this.BtnLoadDictionary.Location = new System.Drawing.Point(405, 43);
            this.BtnLoadDictionary.Name = "BtnLoadDictionary";
            this.BtnLoadDictionary.Size = new System.Drawing.Size(125, 25);
            this.BtnLoadDictionary.TabIndex = 10;
            this.BtnLoadDictionary.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnLoadDictionary.Values.Text = "Load Dictionary";
            this.BtnLoadDictionary.Click += new System.EventHandler(this.BtnLoadDictionary_Click);
            // 
            // BtnSelectionMode
            // 
            this.BtnSelectionMode.Location = new System.Drawing.Point(848, 12);
            this.BtnSelectionMode.Name = "BtnSelectionMode";
            this.BtnSelectionMode.Size = new System.Drawing.Size(125, 25);
            this.BtnSelectionMode.TabIndex = 11;
            this.BtnSelectionMode.Values.DropDownArrowColor = System.Drawing.Color.Empty;
            this.BtnSelectionMode.Values.Text = "Cell Select";
            this.BtnSelectionMode.Click += new System.EventHandler(this.BtnSelectionMode_Click);
            // 
            // kryptonTableLayoutPanel1
            // 
            this.kryptonTableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonTableLayoutPanel1.ColumnCount = 1;
            this.kryptonTableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.kryptonTableLayoutPanel1.Controls.Add(this.outlookGrid1, 0, 0);
            this.kryptonTableLayoutPanel1.Controls.Add(this.kryptonExtraGrid1, 0, 1);
            this.kryptonTableLayoutPanel1.Location = new System.Drawing.Point(12, 74);
            this.kryptonTableLayoutPanel1.Name = "kryptonTableLayoutPanel1";
            this.kryptonTableLayoutPanel1.RowCount = 2;
            this.kryptonTableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.kryptonTableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.kryptonTableLayoutPanel1.Size = new System.Drawing.Size(1005, 552);
            this.kryptonTableLayoutPanel1.TabIndex = 13;
            // 
            // kryptonExtraGrid1
            // 
            this.kryptonExtraGrid1.ButtonSpecs.Add(this.buttonSpecHeaderGroup1);
            this.kryptonExtraGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonExtraGrid1.HeaderVisibleSecondary = false;
            this.kryptonExtraGrid1.Location = new System.Drawing.Point(3, 279);
            // 
            // kryptonExtraGrid1.OutlookGrid
            // 
            this.kryptonExtraGrid1.OutlookGrid.AllowDrop = true;
            this.kryptonExtraGrid1.OutlookGrid.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.kryptonExtraGrid1.OutlookGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonExtraGrid1.OutlookGrid.EnableSearchOnKeyPress = true;
            this.kryptonExtraGrid1.OutlookGrid.Location = new System.Drawing.Point(0, 0);
            this.kryptonExtraGrid1.OutlookGrid.Name = "OutlookGrid";
            this.kryptonExtraGrid1.OutlookGrid.ShowColumnFilter = true;
            this.kryptonExtraGrid1.OutlookGrid.Size = new System.Drawing.Size(997, 238);
            this.kryptonExtraGrid1.OutlookGrid.TabIndex = 0;
            this.kryptonExtraGrid1.ShowGrandTotalAtBottom = true;
            this.kryptonExtraGrid1.ShowGroupBox = false;
            this.kryptonExtraGrid1.ShowSearchToolBar = false;
            this.kryptonExtraGrid1.Size = new System.Drawing.Size(999, 270);
            this.kryptonExtraGrid1.TabIndex = 1;
            this.kryptonExtraGrid1.ValuesPrimary.Image = null;
            this.kryptonExtraGrid1.ValuesSecondary.Heading = "";
            // 
            // buttonSpecHeaderGroup1
            // 
            this.buttonSpecHeaderGroup1.Text = "Hi";
            this.buttonSpecHeaderGroup1.UniqueName = "e788dad1d2104ab487187eac5dacf0ff";
            // 
            // _textBox
            // 
            this._textBox.Location = new System.Drawing.Point(553, 46);
            this._textBox.Name = "_textBox";
            this._textBox.Size = new System.Drawing.Size(100, 23);
            this._textBox.TabIndex = 14;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1029, 638);
            this.Controls.Add(this._textBox);
            this.Controls.Add(this.kryptonTableLayoutPanel1);
            this.Controls.Add(this.BtnSelectionMode);
            this.Controls.Add(this.BtnLoadDictionary);
            this.Controls.Add(this.BtnLoadListOfRawArray);
            this.Controls.Add(this.BtnLoadListOfT);
            this.Controls.Add(this.kryptonButton1);
            this.Controls.Add(this.BtnShowIdInColumnContext);
            this.Controls.Add(this.BtnHighlightSearch);
            this.Controls.Add(this.BtnEnableSearch);
            this.Controls.Add(this.BtnShowColumnFilter);
            this.Controls.Add(this.BtnShowGrandTotal);
            this.Controls.Add(this.BtnShowSubTotal);
            this.Name = "Form1";
            this.Text = "Form1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.outlookGrid1)).EndInit();
            this.kryptonTableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonExtraGrid1.OutlookGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonExtraGrid1.Panel)).EndInit();
            this.kryptonExtraGrid1.Panel.ResumeLayout(false);
            this.kryptonExtraGrid1.Panel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonExtraGrid1)).EndInit();
            this.kryptonExtraGrid1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Krypton.Toolkit.KryptonOutlookGrid outlookGrid1;
        private Krypton.Toolkit.KryptonContextMenu kryptonContextMenu1;
        private Krypton.Toolkit.KryptonContextMenuItem kryptonContextMenuItem1;
        private Krypton.Toolkit.KryptonContextMenuMonthCalendar kryptonContextMenuMonthCalendar1;
        private Krypton.Toolkit.KryptonContextMenuCheckBox kryptonContextMenuCheckBox1;
        private Krypton.Toolkit.KryptonContextMenuHeading kryptonContextMenuHeading1;
        private Krypton.Toolkit.KryptonContextMenuRadioButton kryptonContextMenuRadioButton1;
        private Krypton.Toolkit.KryptonContextMenuCheckButton kryptonContextMenuCheckButton1;
        private Krypton.Toolkit.KryptonContextMenuLinkLabel kryptonContextMenuLinkLabel1;
        private Krypton.Toolkit.KryptonContextMenuColorColumns kryptonContextMenuColorColumns1;
        private Krypton.Toolkit.KryptonContextMenuImageSelect kryptonContextMenuImageSelect1;
        private Krypton.Toolkit.KryptonContextMenuMonthCalendar kryptonContextMenuMonthCalendar2;
        private Krypton.Toolkit.KryptonContextMenuItem kryptonContextMenuItem2;
        private Krypton.Toolkit.KryptonButton BtnShowSubTotal;
        private Krypton.Toolkit.KryptonButton BtnShowGrandTotal;
        private Krypton.Toolkit.KryptonButton BtnShowColumnFilter;
        private Krypton.Toolkit.KryptonButton BtnEnableSearch;
        private Krypton.Toolkit.KryptonButton BtnHighlightSearch;
        private Krypton.Toolkit.KryptonButton BtnShowIdInColumnContext;
        private Krypton.Toolkit.KryptonButton kryptonButton1;
        private Krypton.Toolkit.KryptonButton BtnLoadListOfT;
        private Krypton.Toolkit.KryptonButton BtnLoadListOfRawArray;
        private Krypton.Toolkit.KryptonButton BtnLoadDictionary;
        private Krypton.Toolkit.KryptonButton BtnSelectionMode;
        private Krypton.Toolkit.KryptonTableLayoutPanel kryptonTableLayoutPanel1;
        private Krypton.Toolkit.KryptonTextBox _textBox;
        private Krypton.Toolkit.KryptonAllInOneGrid kryptonExtraGrid1;
        private Krypton.Toolkit.ButtonSpecHeaderGroup buttonSpecHeaderGroup1;
    }
}
