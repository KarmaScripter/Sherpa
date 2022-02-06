﻿// <copyright file = "ImageList.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>

namespace BudgetExecution
{
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Drawing;
    using System.Windows.Forms;
    using Syncfusion.Windows.Forms.Tools;

    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "AutoPropertyCanBeMadeGetOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "AutoPropertyCanBeMadeGetOnly.Global" ) ]
    public class ImageList : ImageListAdv
    {
        /// <summary>
        /// Initializes a new instance of the <see cref = "ImageList"/> class.
        /// </summary>
        public ImageList()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "ImageList"/> class.
        /// </summary>
        /// <param name = "imageSource" >
        /// The image source.
        /// </param>
        /// <param name = "size" >
        /// The size.
        /// </param>
        public ImageList( ImageSource imageSource, Size size )
        {
            Source = imageSource;
            ImageSize = size;
            Builder = new ImageBuilder( Source );
            Factory = new ImageFactory( Builder );
            UseImageSize = true;
        }
        
        /// <summary>
        /// Gets or sets the image builder.
        /// </summary>
        /// <value>
        /// The image builder.
        /// </value>
        private ImageBuilder Builder { get; }

        /// <summary>
        /// Gets or sets the image factory.
        /// </summary>
        /// <value>
        /// The image factory.
        /// </value>
        private ImageFactory Factory { get; }

        /// <summary>
        /// Gets or sets the budget images.
        /// </summary>
        /// <value>
        /// The budget images.
        /// </value>
        public IEnumerable<BudgetImage> BudgetImages { get; set; }

        /// <summary>
        /// Gets or sets the image source.
        /// </summary>
        /// <value>
        /// The image source.
        /// </value>
        public ImageSource Source { get; set; }

        /// <summary>
        /// Gets or sets the binding source.
        /// </summary>
        /// <value>
        /// The binding souce.
        /// </value>
        public BindingSource BindingSource { get; set; }

        /// <summary>
        /// Gets or sets the image.
        /// </summary>
        /// <value>
        /// The image.
        /// </value>
        public BudgetImage BudgetImage { get; set; }
    }
}