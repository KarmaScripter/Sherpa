﻿// ******************************************************************************************
//     Assembly:             BudgetExecution
//     Author:                  Terry D. Eppler
//     Created:                 06-19-2023
// 
//     Last Modified By:        Terry D. Eppler
//     Last Modified On:        11-25-2023
// ******************************************************************************************
// <copyright file="Terry Eppler.cs" company="Terry D. Eppler">
//    BudgetExecution is a Federal Budget, Finance, and Accounting application for the
//    US Environmental Protection Agency (US EPA).
//    Copyright ©  2023  Terry Eppler
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
//    You can contact me at:  terryeppler@gmail.com or eppler.terry@epa.gov
// </copyright>
// <summary>
//   ListView.cs.cs
// </summary>
// ******************************************************************************************

namespace BudgetExecution
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing;
    using System.Windows.Forms;
    using Syncfusion.WinForms.ListView;

    /// <inheritdoc />
    /// <summary> </summary>
    [ SuppressMessage( "ReSharper", "ClassCanBeSealed.Global" ) ]
    public class ListView : SfListView
    {
        /// <summary> Gets or sets the binding source. </summary>
        /// <value> The binding source. </value>
        public BindingSource BindingSource { get; set; }

        /// <summary> Gets or sets the tool tip. </summary>
        /// <value> The tool tip. </value>
        public SmallTip ToolTip { get; set; }

        /// <summary> Gets or sets the hover text. </summary>
        /// <value> The hover text. </value>
        public string HoverText { get; set; }

        /// <summary> Gets or sets the filter. </summary>
        /// <value> The filter. </value>
        public IDictionary<string, object> DataFilter { get; set; }

        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="ListView"/>
        /// class.
        /// </summary>
        public ListView( )
        {
            // Basic Properties
            Size = new Size( 250, 150 );
            Font = new Font( "Roboto", 8 );
            Margin = new Padding( 1 );
            Padding = new Padding( 1 );
            Anchor = AnchorStyles.Top | AnchorStyles.Left;
            Dock = DockStyle.None;
            Enabled = true;
            Visible = true;

            // BackColor SeriesConfiguration
            BackColor = Color.FromArgb( 20, 20, 20 );
            ForeColor = Color.FromArgb( 106, 189, 252 );

            // Border SeriesConfiguration
            CanOverrideStyle = true;
            HotTracking = true;

            // Item SeriesConfiguration
            ItemHeight = 24;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.ListView" />
        /// class.
        /// </summary>
        /// <param name="size"> The size. </param>
        /// <param name="location"> The location. </param>
        public ListView( Size size, Point location )
            : this( )
        {
            Size = size;
            Location = location;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.ListView" />
        /// class.
        /// </summary>
        /// <param name="label"> The label. </param>
        public ListView( Label label )
            : this( label.Size, label.Location )
        {
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.ListView" />
        /// class.
        /// </summary>
        /// <param name="size"> The size. </param>
        /// <param name="location"> The location. </param>
        /// <param name="parent"> The parent. </param>
        public ListView( Size size, Point location, Control parent )
            : this( size, location )
        {
            Parent = parent;
            Parent.Controls.Add( this );
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.ListView" />
        /// class.
        /// </summary>
        /// <param name="size"> The size. </param>
        /// <param name="location"> The location. </param>
        /// <param name="parent"> The parent. </param>
        /// <param name="text"> The text. </param>
        public ListView( Size size, Point location, Control parent, string text )
            : this( size, location, parent )
        {
            Text = text;
        }

        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="ListView"/>
        /// class.
        /// </summary>
        /// <param name="size"> The size. </param>
        /// <param name="location"> The location. </param>
        /// <param name="parent"> The parent. </param>
        /// <param name="bindingSource"> The binding source. </param>
        public ListView( Size size, Point location, Control parent, BindingSource bindingSource )
            : this( size, location, parent )
        {
            BindingSource = bindingSource;
        }

        /// <summary>
        /// Sets the hover text.
        /// </summary>
        /// <param name="text">The text.</param>
        public void SetHoverText( string text )
        {
            if( !string.IsNullOrEmpty( text ) )
            {
                try
                {
                    var _ = new SmallTip( this, text );
                }
                catch( Exception _ex )
                {
                    Fail( _ex );
                }
            }
        }

        /// <summary>
        /// Sets the text.
        /// </summary>
        /// <param name="text">The text.</param>
        public void SetText( string text )
        {
            if( !string.IsNullOrEmpty( text ) )
            {
                try
                {
                    Text = text;
                }
                catch( Exception _ex )
                {
                    Fail( _ex );
                }
            }
        }

        /// <summary>
        /// Fails the specified ex.
        /// </summary>
        /// <param name="ex">The ex.</param>
        private void Fail( Exception ex )
        {
            using var _error = new ErrorDialog( ex );
            _error?.SetText( );
            _error?.ShowDialog( );
        }
    }
}