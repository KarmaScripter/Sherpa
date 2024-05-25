﻿//  ******************************************************************************************
//      Assembly:                Sherpa
//      Filename:                RibbonBase.cs
//      Author:                  Terry D. Eppler
//      Created:                 05-31-2023
// 
//      Last Modified By:        Terry D. Eppler
//      Last Modified On:        06-01-2023
//  ******************************************************************************************
//  <copyright file="RibbonBase.cs" company="Terry D. Eppler">
// 
//     This is a Federal Budget, Finance, and Accounting application for the
//     US Environmental Protection Agency (US EPA).
//     Copyright ©  2023  Terry Eppler
// 
//     Permission is hereby granted, free of charge, to any person obtaining a copy
//     of this software and associated documentation files (the “Software”),
//     to deal in the Software without restriction,
//     including without limitation the rights to use,
//     copy, modify, merge, publish, distribute, sublicense,
//     and/or sell copies of the Software,
//     and to permit persons to whom the Software is furnished to do so,
//     subject to the following conditions:
// 
//     The above copyright notice and this permission notice shall be included in all
//     copies or substantial portions of the Software.
// 
//     THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
//     INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//     FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
//     IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
//     DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
//     ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
//     DEALINGS IN THE SOFTWARE.
// 
//     You can contact me at:   terryeppler@gmail.com or eppler.terry@epa.gov
// 
//  </copyright>
//  <summary>
//    RibbonBase.cs
//  </summary>
//  ******************************************************************************************

namespace Sherpa
{
    using System;
    using System.Drawing;
    using System.Windows.Forms;
    using Syncfusion.Windows.Forms.Spreadsheet;
    using Syncfusion.Windows.Forms.Tools;

    public class RibbonBase : SpreadsheetRibbon
    {
        /// <summary> Gets or sets the grid. </summary>
        /// <value> The grid. </value>
        public virtual SpreadsheetGrid Grid { get; set; }

        /// <summary> Gets or sets the active sheet. </summary>
        /// <value> The active sheet. </value>
        public virtual Spreadsheet ActiveSheet { get; set; }

        /// <summary> Gets or sets the model. </summary>
        /// <value> The model. </value>
        public virtual SpreadsheetGridModel Model { get; set; }

        /// <summary> Gets or sets the binding source. </summary>
        /// <value> The binding source. </value>
        public virtual BindingSource BindingSource { get; set; }

        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="RibbonBase"/>
        /// class.
        /// </summary>
        public RibbonBase( )
        {
            EnableRibbonCustomization = true;
            Margin = new Padding( 3 );
            Padding = new Padding( 1 );
            Font = new Font( "Roboto", 9 );
            ForeColor = Color.Black;
            BackColor = Color.FromArgb( 20, 20, 20 );
            BorderStyle = ToolStripBorderStyle.None;
            RibbonStyle = RibbonStyle.Office2010;
            OfficeColorScheme = ToolStripEx.ColorScheme.Black;
            TitleFont = new Font( "Roboto", 9 );

            // Office Menu Properties
            OfficeMenu.BackColor = Color.FromArgb( 20, 20, 20 );
            OfficeMenu.Font = new Font( "Roboto", 9 );
            OfficeMenu.AutoSize = true;
            OfficeMenu.LayoutStyle = ToolStripLayoutStyle.Flow;
            ShowQuickItemsDropDownButton = false;
            Ribbon.ScaleMenuButtonImage = true;
        }

        /// <summary> Fails the specified ex. </summary>
        /// <param name="ex"> The ex. </param>
        private void Fail( Exception ex )
        {
            using var _error = new ErrorDialog( ex );
            _error?.SetText( );
            _error?.ShowDialog( );
        }
    }
}