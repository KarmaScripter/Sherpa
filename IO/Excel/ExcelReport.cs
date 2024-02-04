﻿// ******************************************************************************************
//     Assembly:                Budget Execution
//     Author:                  Terry D. Eppler
//     Created:                 2-3-2024
// 
//     Last Modified By:        Terry D. Eppler
//     Last Modified On:        2-3-2024
// ******************************************************************************************
// <copyright file="ExcelReport.cs" company="Terry D. Eppler">
//    Budget Execution is a Federal Budget, Finance, and Accounting application
//    for analysts with the US Environmental Protection Agency (US EPA).
//    Copyright ©  2024  Terry Eppler
// 
//    Permission is hereby granted, free of charge, to any person obtaining a copy
//    of this software and associated documentation files (the “Software”),
//    to deal in the Software without restriction,
//    including without limitation the rights to use,
//    copy, modify, merge, publish, distribute, sublicense,
//    and/or sell copies of the Software,
//    and to permit persons to whom the Software is furnished to do so,
//    subject to the following conditions:
// 
//    The above copyright notice and this permission notice shall be included in all
//    copies or substantial portions of the Software.
// 
//    THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
//    INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//    FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
//    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
//    DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
//    ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
//    DEALINGS IN THE SOFTWARE.
// 
//    Contact at:  terryeppler@gmail.com or eppler.terry@epa.gov
// </copyright>
// <summary>
//   ExcelReport.cs
// </summary>
// ******************************************************************************************

namespace BudgetExecution
{
    using System;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing;
    using System.IO;
    using OfficeOpenXml;
    using OfficeOpenXml.Style;
    using static System.Configuration.ConfigurationManager;

    /// <inheritdoc/>
    /// <summary> </summary>
    /// <seealso cref="T:BudgetExecution.SheetConfig"/>
    [ SuppressMessage( "ReSharper", "MergeIntoPattern" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "UnusedParameter.Global" ) ]
    [ SuppressMessage( "ReSharper", "UnusedType.Global" ) ]
    [ SuppressMessage( "ReSharper", "AutoPropertyCanBeMadeGetOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "AssignNullToNotNullAttribute" ) ]
    [ SuppressMessage( "ReSharper", "PropertyCanBeMadeInitOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeInternal" ) ]
    [ SuppressMessage( "ReSharper", "RedundantCheckBeforeAssignment" ) ]
    [ SuppressMessage( "ReSharper", "ClassCanBeSealed.Global" ) ]
    [ SuppressMessage( "ReSharper", "PossibleNullReferenceException" ) ]
    public class ExcelReport : PartFactory
    {
        /// <summary>
        /// Gets or sets the height of the row.
        /// </summary>
        /// <value>
        /// The height of the row.
        /// </value>
        public double RowHeight
        {
            get
            {
                return _rowHeight;
            }

            private protected set
            {
                _rowHeight = value;
            }
        }

        /// <summary>
        /// Gets or sets the width of the column.
        /// </summary>
        /// <value>
        /// The width of the column.
        /// </value>
        public double ColumnWidth
        {
            get
            {
                return _columnWidth;
            }

            private protected set
            {
                _columnWidth = value;
            }
        }

        /// <summary>
        /// The configuration
        /// </summary>
        public string InternalPath
        {
            get
            {
                return _internalPath;
            }

            private protected set
            {
                _internalPath = value;
            }
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.ExcelReport" /> class.
        /// </summary>
        public ExcelReport( )
        {
            _rowIndex = 2;
            _fontColor = Color.Black;
            _font = new Font( "Roboto", 8, FontStyle.Regular );
            _titleFont = new Font( "Roboto", 9 );
            _fileName = "Budget.xlsx";
            _rowHeight = 0.27d;
            _columnWidth = 0.69d;
            _leftMargin = .25m;
            _rightMargin = .25m;
            _topMargin = 1m;
            _bottomMargin = 1m;
            _headerMargin = 0.4m;
            _footerMargin = 0.4m;
            _zoomLevel = 100;
            _primaryBackColor = Color.White;
            _secondaryBackColor = Color.FromArgb( 220, 220, 220 );
            _rowCount = 55;
            _columnCount = 11;
            _internalPath = AppSettings[ "Reports" ];
            _savePath = AppSettings[ "Desktop" ] + _fileName;
            _fileInfo = new FileInfo( _internalPath );
            _excelPackage = new ExcelPackage( _fileInfo );
            _excelPackage.Settings.TextSettings.PrimaryTextMeasurer = new TextSize( );
            _excelPackage.Settings.TextSettings.AutofitScaleFactor = 0.9f;
            _excelWorkbook = _excelPackage.Workbook;
            _dataWorksheet = _excelWorkbook.Worksheets.Add( "Data" );
            _dataRange = _dataWorksheet.Cells[ 2, 1, 57, 11 ];
            _excelWorkbook.View.ShowHorizontalScrollBar = true;
            _excelWorkbook.View.ShowVerticalScrollBar = true;
            InitializeWorkbook( );
            InitializeActiveGrid( );
            InitializeSheetView( );
            InitializePrinterSettings( );
            InitializeHeaderFooter( );
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetWorkbook" /> class.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public ExcelReport( string filePath )
            : this( )
        {
            _rowIndex = 2;
            _fontColor = Color.Black;
            _font = new Font( "Roboto", 8, FontStyle.Regular );
            _titleFont = new Font( "Roboto", 9 );
            _rowHeight = 0.27d;
            _columnWidth = 0.69d;
            _leftMargin = .25m;
            _rightMargin = .25m;
            _topMargin = 1m;
            _bottomMargin = 1m;
            _headerMargin = 0.4m;
            _footerMargin = 0.4m;
            _zoomLevel = 100;
            _primaryBackColor = Color.White;
            _secondaryBackColor = Color.FromArgb( 220, 220, 220 );
            _internalPath = AppSettings[ "Reports" ];
            _savePath = AppSettings[ "Desktop" ] + _fileName + ".xlsx";
            _filePath = filePath;
            _fileName = Path.GetFileNameWithoutExtension( filePath );
            _fileInfo = new FileInfo( filePath );
            _excelPackage = new ExcelPackage( _fileInfo );
            _excelPackage.Settings.TextSettings.PrimaryTextMeasurer = new TextSize( );
            _excelPackage.Settings.TextSettings.AutofitScaleFactor = 0.9f;
            _excelWorkbook = _excelPackage.Workbook;
            _dataWorksheet = _excelWorkbook.Worksheets[ 0 ];
            _excelWorkbook.View.ShowHorizontalScrollBar = true;
            _excelWorkbook.View.ShowVerticalScrollBar = true;
            InitializeWorkbook( );
            InitializeActiveGrid( );
            InitializeSheetView( );
            InitializePrinterSettings( );
            InitializeHeaderFooter( );
        }

        /// <inheritdoc/>
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetWorkbook"/>
        /// class.
        /// </summary>
        /// <param name="dataTable">
        /// The data table.
        /// </param>
        public ExcelReport( DataTable dataTable )
        {
            _rowIndex = 2;
            _fontColor = Color.Black;
            _font = new Font( "Roboto", 8, FontStyle.Regular );
            _titleFont = new Font( "Roboto", 9, FontStyle.Regular );
            _rowHeight = 0.27d;
            _columnWidth = 0.69d;
            _leftMargin = .25m;
            _rightMargin = .25m;
            _topMargin = 1m;
            _bottomMargin = 1m;
            _headerMargin = 0.4m;
            _footerMargin = 0.4m;
            _zoomLevel = 100;
            _primaryBackColor = Color.White;
            _secondaryBackColor = Color.FromArgb( 220, 220, 220 );
            _internalPath = AppSettings[ "Reports" ];
            _dataTable = dataTable;
            _columnCount = dataTable.Columns.Count;
            _rowCount = dataTable.Rows.Count;
            _fileName = dataTable.TableName + ".xlsx";
            _savePath = AppSettings[ "Desktop" ] + dataTable.TableName + ".xlsx";
            _dataMetric = new DataMetric( dataTable );
            _fileInfo = new FileInfo( _internalPath );
            _excelPackage = new ExcelPackage( _fileInfo );
            _excelPackage.Settings.TextSettings.PrimaryTextMeasurer = new TextSize( );
            _excelPackage.Settings.TextSettings.AutofitScaleFactor = 0.8f;
            _excelWorkbook = _excelPackage.Workbook;
            _dataWorksheet = _excelWorkbook.Worksheets.Add( "Data" );
            _excelWorkbook.View.ShowHorizontalScrollBar = true;
            _excelWorkbook.View.ShowVerticalScrollBar = true;
            InitializeWorkbook( );
            InitializeActiveGrid( );
            InitializeSheetView( );
            InitializePrinterSettings( );
            InitializeHeaderFooter( );
        }

        /// <summary>
        /// Initializes the worksheets.
        /// </summary>
        private void InitializeSheetView( )
        {
            try
            {
                for( var _i = 0; _i < _excelWorkbook.Worksheets.Count; _i++ )
                {
                    _excelWorkbook.Worksheets[ _i ].View.ShowGridLines = false;
                    _excelWorkbook.Worksheets[ _i ].View.PageLayoutView = true;
                    _excelWorkbook.Worksheets[ _i ].View.ShowHeaders = true;
                    _excelWorkbook.Worksheets[ _i ].View.ZoomScale = _zoomLevel;
                }
            }
            catch( Exception _ex )
            {
                Fail( _ex );
            }
        }

        /// <summary>
        /// Initializes the printer settings.
        /// </summary>
        private void InitializePrinterSettings( )
        {
            try
            {
                for( var _i = 0; _i < _excelWorkbook.Worksheets.Count; _i++ )
                {
                    _excelWorkbook.Worksheets[ _i ].PrinterSettings.LeftMargin = _leftMargin;
                    _excelWorkbook.Worksheets[ _i ].PrinterSettings.RightMargin = _rightMargin;
                    _excelWorkbook.Worksheets[ _i ].PrinterSettings.HeaderMargin = _headerMargin;
                    _excelWorkbook.Worksheets[ _i ].PrinterSettings.FooterMargin = _footerMargin;
                    _excelWorkbook.Worksheets[ _i ].PrinterSettings.TopMargin = _topMargin;
                    _excelWorkbook.Worksheets[ _i ].PrinterSettings.BottomMargin = _bottomMargin;
                    _excelWorkbook.Worksheets[ _i ].PrinterSettings.Orientation =
                        eOrientation.Landscape;
                }
            }
            catch( Exception _ex )
            {
                Fail( _ex );
            }
        }

        /// <summary>
        /// Initializes the header footer.
        /// </summary>
        private void InitializeHeaderFooter( )
        {
            try
            {
                var _rows = _dataTable.Rows.Count.ToString( "N0" );
                var _cols = _dataTable.Columns.Count.ToString( );
                _dataWorksheet.HeaderFooter.OddFooter.RightAlignedText =
                    "Records:" + '\t'
                    + "  "
                    + _rows.PadLeft( 16 )
                    + Environment.NewLine
                    + "Columns:" + '\t'
                    + "  "
                    + _cols.PadLeft( 19 );

                _dataWorksheet.HeaderFooter.OddFooter.CenteredText =
                    "Page" + '\r' + '\n' +
                    ExcelHeaderFooter.PageNumber
                    + " of "
                    + ExcelHeaderFooter.NumberOfPages;

                _dataWorksheet.HeaderFooter.OddFooter.LeftAlignedText =
                    "Data as of:"
                    + Environment.NewLine
                    + DateTime.Today.ToLongDateString( );
            }
            catch( Exception _ex )
            {
                Fail( _ex );
            }
        }

        /// <summary>
        /// Initializes the range cells.
        /// </summary>
        private void InitializeActiveGrid( )
        {
            try
            {
                if( _dataTable == null )
                {
                    _dataRange = _dataWorksheet.Cells[ 2, 1, 57, 11 ];
                    _dataRange.Style.Font.Name = "Roboto";
                    _dataRange.Style.Font.Size = 8;
                    _dataRange.Style.Font.Bold = false;
                    _dataRange.Style.Font.Italic = false;
                    _dataRange.EntireRow.CustomHeight = true;
                    _dataRange.Style.Font.Color.SetColor( _fontColor );
                    _dataRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    _dataRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    FormatHeader( _dataRange );
                    SetAlternatingRowColor( _dataRange );
                    FormatFooter( _dataRange );
                }
                else
                {
                    _excelTable = CreateExcelTable( _dataTable );
                    var _fields = _dataMetric?.Fields;
                    var _numerics = _dataMetric?.Numerics;
                    _pivotTable = CreatePivotTable( _dataRange, "PivotData", _fields, _numerics );
                }
            }
            catch( Exception _ex )
            {
                if( _dataRange != null )
                {
                    _dataRange = null;
                }

                Fail( _ex );
            }
        }

        /// <summary>
        /// Initializes the workbook properties.
        /// </summary>
        private void InitializeWorkbook( )
        {
            try
            {
                _excelWorkbook.Properties.Author = "Terry D. Eppler, PhD.";
                _excelWorkbook.Properties.Application = "Budget Execution";
                _excelWorkbook.Properties.Company = "US EPA";
                _excelWorkbook.Properties.Modified = DateTime.Now;
            }
            catch( Exception _ex )
            {
                Fail( _ex );
            }
        }

        /// <summary>
        /// Saves this instance.
        /// </summary>
        public void Save( )
        {
            try
            {
                _excelPackage.SaveAs( _savePath );
            }
            catch( Exception _ex )
            {
                Fail( _ex );
            }
        }
    }
}