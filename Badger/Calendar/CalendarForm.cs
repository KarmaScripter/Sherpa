﻿// <copyright file = "CalendarForm.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>
//

namespace BudgetExecution
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using Syncfusion.Windows.Forms;

    public partial class CalendarForm : MetroForm
    {
        public CalendarPanel Calendar { get; } = new CalendarPanel();

        public CalendarForm()
        {
            InitializeComponent();
            BackColor = ColorConfig.FormBackColorDark;
            BorderThickness = BorderConfig.Thin;
            BorderColor = ColorConfig.BorderColorBlue;
            Font = FontConfig.FontSizeSmall;
            CaptionBarColor = ColorConfig.FormBackColorDark;
            CaptionBarHeight = SizeConfig.CaptionSizeNormal;
            CaptionButtonColor = ColorConfig.CaptionButtonDefaultColor;
            CaptionButtonHoverColor = ColorConfig.CaptionButtonHoverColor;
            CaptionAlign = AlignConfig.HorizontalLeft;
            CaptionFont = FontConfig.FontSizeMedium;
            MetroColor = ColorConfig.FormBackColorDark;
            FormBorderStyle = BorderConfig.Sizeable;
            Icon = new Icon( @"C:\Users\terry\source\repos\BudgetExecution\Etc\epa.ico", 33, 32 );
            ShowIcon = false;
            ShowInTaskbar = true;
            Padding = ControlConfig.Padding;
            Text = string.Empty;
            MinimumSize = SizeConfig.FormSizeMinimum;
            MaximumSize = SizeConfig.FormSizeMaximum;
            Calendar.Dock = DockStyle.Fill;
            Size = new Size( Calendar.Size.Width + 5, Calendar.Size.Height + 5 );
            Controls.Add( Calendar );
        }
    }
}
