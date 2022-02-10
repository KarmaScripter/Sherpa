﻿// <copyright file = "BudgetImage.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>

namespace BudgetExecution
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing;
    using System.IO;
    using System.Linq;

    /// <summary>
    /// 
    /// </summary>
    /// <seealso cref="BudgetExecution.BudgetSize" />
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeInternal" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeProtected.Global" ) ]
    [ SuppressMessage( "ReSharper", "VirtualMemberNeverOverridden.Global" ) ]
    public abstract class ImageBase : BudgetSize 
    {
        /// <summary>
        /// Gets or sets the image.
        /// </summary>
        /// <value>
        /// The image.
        /// </value>
        public virtual Bitmap ImageFile { get; set; }

        /// <summary>
        /// Gets or sets the source.
        /// </summary>
        /// <value>
        /// The source.
        /// </value>
        public virtual ImageSource ImageSource { get; set; }

        /// <summary>
        /// Gets or sets the format.
        /// </summary>
        /// <value>
        /// The format.
        /// </value>
        public virtual ImageFormat ImageFormat { get; set; }

        /// <summary>
        /// Gets or sets the color of the back ground.
        /// </summary>
        /// <value>
        /// The color of the back ground.
        /// </value>
        public virtual Color BackGroundColor { get; set; } = Color.Transparent;

        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        /// <value>
        /// The color.
        /// </value>
        public virtual Color Color { get; set; }

        /// <summary>
        /// Gets or sets the size.
        /// </summary>
        /// <value>
        /// The size.
        /// </value>
        public virtual Size Size { get; set; }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public virtual string Name { get; set; }

        /// <summary>
        /// Gets or sets the file extension.
        /// </summary>
        /// <value>
        /// The file extension.
        /// </value>
        public virtual string FileExtension { get; set; }

        /// <summary>
        /// Gets the image source.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <returns></returns>
        public virtual ImageSource GetImageSource( ImageSource source )
        {
            try
            {
                return Verify.ImageResource( source )
                    ? source
                    : ImageSource.NS;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return ImageSource.NS;
            }
        }

        /// <summary>
        /// Gets the file extension.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="resource">The resource.</param>
        /// <returns></returns>
        public virtual string GetFileExtension( string filePath, ImageSource resource = ImageSource.NS )
        {
            try
            {
                return Verify.ImageResource( resource )
                    && Verify.IsInput( filePath )
                    && File.Exists( filePath )
                    && resource != ImageSource.NS
                        ? filePath
                        : string.Empty;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the image format.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <returns></returns>
        public virtual ImageFormat GetImageFormat( ImageFormat format )
        {
            try
            {
                return Enum.IsDefined( typeof( ImageFormat ), format )
                    ? format
                    : ImageFormat.NS;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return ImageFormat.NS;
            }
        }

        /// <summary>
        /// Gets the image format.
        /// </summary>
        /// <param name="extension">The extension.</param>
        /// <returns></returns>
        public virtual ImageFormat GetImageFormat( string extension )
        {
            try
            {
                var _names = Enum.GetNames( typeof( ImageFormat ) );

                return _names.Contains( extension )
                    ? (ImageFormat)Enum.Parse( typeof( ImageFormat ), extension )
                    : ImageFormat.PNG;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return ImageFormat.NS;
            }
        }

        /// <summary>
        /// Gets the size of the image.
        /// </summary>
        /// <param name="size">The size.</param>
        /// <returns></returns>
        public virtual Size GetImageSize( Size size )
        {
            try
            {
                return size != Size.Empty
                    ? GetSize( size )
                    : Size.Empty;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Size.Empty;
            }
        }

        /// <summary>
        /// Gets the size of the image.
        /// </summary>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <returns></returns>
        public virtual Size GetImageSize( int width, int height )
        {
            try
            {
                return width > -1 && height > -1
                    ? GetSize( width, height )
                    : Size.Empty;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Size.Empty;
            }
        }

        /// <summary>
        /// Gets the size of the image.
        /// </summary>
        /// <param name="sizer">The sizer.</param>
        /// <returns></returns>
        public virtual Size GetImageSize( ImageSizer sizer )
        {
            try
            {
                return Enum.IsDefined( typeof( ImageSizer ), sizer )
                    ? GetSize( sizer )
                    : Size.Empty;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Size.Empty;
            }
        }

        /// <summary>
        /// Sets the color of the back ground.
        /// </summary>
        /// <param name="newColor">The new color.</param>
        public virtual void SetBackGroundColor( Color newColor )
        {
            if( newColor != Color.Empty )
            {
                try
                {
                    BackGroundColor = newColor;
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }
        }
    }
}