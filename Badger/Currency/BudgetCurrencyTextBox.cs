﻿// <copyright file = "BudgetCurrencyTextBox.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>
//

namespace BudgetExecution
{
    using System;
    using System.Windows.Forms;
    using Syncfusion.Windows.Forms.Tools;

    public class BudgetCurrencyTextBox : CurrencyBoxBase
    {
        /// <summary>
        /// Initializes a new instance of
        /// the <see cref="BudgetCurrencyTextBox"/> class.
        /// </summary>
        /// <remarks>
        /// The CurrencyEdit class also creates
        /// the controls that it hosts such
        /// as the
        /// <see cref="T:Syncfusion.Windows.Forms.Tools.CurrencyTextBox" />
        /// control and the
        /// <see cref="T:Syncfusion.Windows.Forms.Tools.PopupCalculator" />
        /// control.
        /// </remarks>
        public BudgetCurrencyTextBox()
        {
            // Basic Properties
            Size = BudgetSize.TextBoxControl;
            Location = BudgetSetting.GetLocation();
            Anchor = BudgetSetting.GetAnchorStyle();
            Dock = DockStyle.None;
            Margin = BudgetSetting.Margin;
            Padding = BudgetSetting.Padding;
            Font = BudgetFont.FontSizeSmall;
            ForeColor = BudgetColor.LightGray;
            Enabled = true;
            Visible = true;
            Text = string.Empty;
            Border3DStyle = Border3DStyle.Flat;
            PopupCalculatorAlignment = CalculatorPopupAlignment.Right;
            FlatStyle = FlatStyle.Flat;
            ShowCalculator = true;
            TextAlign = HorizontalAlignment.Center;

            // TextBox Properties
            TextBox.CurrencyDecimalDigits = 2;
            TextBox.NegativeColor = BudgetColor.Red;
            TextBox.PositiveColor = BudgetColor.LightBlue;
            TextBox.BackGroundColor = BudgetColor.FormDark;
            TextBox.Border3DStyle = Border3DStyle.Flat;
            TextBox.ThemeStyle = CurrencyTextBoxVisualStyle.DefaultStyle;
            TextBox.BorderColor = BudgetColor.BorderDark;
            TextBox.BorderStyle = BorderStyle.Fixed3D;
            TextBox.CurrencyDecimalSeparator = ".";
            TextBox.CurrencyGroupSeparator = ",";
            TextBox.FocusBorderColor = BudgetColor.SteelBlue;
            TransferFromCalculator = true;
            TransferToCalculator = true;

            // TextBox ThemeStyle Properties
            ThemeStyle.BackColor = BudgetColor.FormDark;
            ThemeStyle.BorderColor = BudgetColor.BorderDark;
            ThemeStyle.DisabledBackColor = BudgetColor.FormDark;
            ThemeStyle.DisabledBorderColor  = BudgetColor.FormDark;
            ThemeStyle.HoverBorderColor = BudgetColor.SteelBlue;
            ThemeStyle.FocusedBorderColor = BudgetColor.SteelBlue;
            ThemeStyle.PressedBorderColor = BudgetColor.SteelBlue;
        }
    }
}
