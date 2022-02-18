﻿// <copyright file = "BudgetChart.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>

namespace BudgetExecution
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing;
    using System.Drawing.Drawing2D;
    using System.Windows.Forms;
    using Syncfusion.Drawing;
    using Syncfusion.Windows.Forms.Chart;

    /// <summary>
    /// 
    /// </summary>
    /// <seealso cref="ChartStyleInfo" />
    /// <seealso cref="IChartConfig" />
    [ SuppressMessage( "ReSharper", "AutoPropertyCanBeMadeGetOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberInitializerValueIgnored" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeInternal" ) ]
    [ SuppressMessage( "ReSharper", "UseObjectOrCollectionInitializer" ) ]
    public class ChartConfig : ChartStyleInfo, IChartConfig
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ChartConfig"/> class.
        /// </summary>
        public ChartConfig()
        {
            DisplayText = true;
            TextColor = Color.FromArgb( 141, 139, 138 );
            Interior = new BrushInfo( Color.FromArgb( 10, 10, 10 ) );
            Border = SetBorderStyle( Color.FromArgb( 10, 10, 10 ), 1 );
        }

        /// <summary>
        /// Sets the font.
        /// </summary>
        /// <param name="font">The font.</param>
        /// <returns></returns>
        public static ChartFontInfo SetFont( Font font )
        {
            if( font != null )
            {
                try
                {
                    var _chart = new ChartFontInfo
                    {
                        FontStyle = font.Style,
                        FontFamilyTemplate = font.FontFamily,
                        Size = font.Size
                    };

                    return _chart;
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }

            return default( ChartFontInfo );
        }

        /// <summary>
        /// Sets the font.
        /// </summary>
        /// <param name="family">The family.</param>
        /// <param name="size">The size.</param>
        /// <param name="style">The style.</param>
        /// <returns></returns>
        public static ChartFontInfo SetFont( string family = "Roboto", int size = 8,
            FontStyle style = FontStyle.Regular )
        {
            if( Verify.IsInput( family ) )
            {
                try
                {
                    var _chart = new ChartFontInfo
                    {
                        FontStyle = style,
                        Facename = family,
                        Size = size
                    };

                    return _chart;
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }

            return default( ChartFontInfo );
        }

        /// <summary>
        /// Sets the anchor.
        /// </summary>
        /// <param name="anchor">The anchor.</param>
        /// <returns></returns>
        public AnchorStyles SetAnchor( AnchorStyles anchor = AnchorStyles.Left & AnchorStyles.Top )
        {
            try
            {
                return anchor;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( AnchorStyles );
            }
        }

        /// <summary>
        /// Sets the border style.
        /// </summary>
        /// <param name="color">The color.</param>
        /// <param name="width">The width.</param>
        /// <returns></returns>
        public static ChartLineInfo SetBorderStyle( Color color, int width = 1 )
        {
            try
            {
                var _style = new ChartLineInfo
                {
                    Alignment = PenAlignment.Center,
                    Width = width,
                    Color = color
                };

                return _style;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( ChartLineInfo );
            }
        }

        /// <summary>
        /// Fails the specified ex.
        /// </summary>
        /// <param name="ex">The ex.</param>
        private protected static void Fail( Exception ex )
        {
            using var _error = new Error( ex );
            _error?.SetText();
            _error?.ShowDialog();
        }
    }
}