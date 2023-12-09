﻿// ******************************************************************************************
//     Assembly:                Budget Execution
//     Author:                  Terry D. Eppler
//     Created:                 05-04-2023
// 
//     Last Modified By:        Terry D. Eppler
//     Last Modified On:        05-31-2023
// ******************************************************************************************
// <copyright file="AnnualReimbursableCarryoverEstimate.cs" company="Terry D. Eppler">
//    This is a Federal Budget, Finance, and Accounting application for the
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
//    You can contact me at:   terryeppler@gmail.com or eppler.terry@epa.gov
// </copyright>
// <summary>
//   AnnualReimbursableCarryoverEstimate.cs
// </summary>
// ******************************************************************************************

namespace BudgetExecution
{
    using System;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;

    /// <inheritdoc/>
    /// <summary> </summary>
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "UnusedType.Global" ) ]
    [ SuppressMessage( "ReSharper", "RedundantBaseConstructorCall" ) ]
    public class AnnualReimbursableCarryoverEstimates : AnnualCarryoverEstimates
    {
        /// <inheritdoc/>
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.AnnualReimbursableCarryoverEstimate"/>
        /// class.
        /// </summary>
        public AnnualReimbursableCarryoverEstimates( )
            : base( )
        {
            Source = Source.AnnualReimbursableEstimates;
        }

        /// <inheritdoc/>
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.AnnualReimbursableCarryoverEstimate"/>
        /// class.
        /// </summary>
        /// <param name="query"> The query. </param>
        public AnnualReimbursableCarryoverEstimates( IQuery query )
            : this( )
        {
            Record = new DataBuilder( query ).Record;
            Data = Record.ToDictionary( );
            ID = int.Parse( Record[ "AnnualReimbursableEstimatesId" ].ToString( ) ?? "0" );
            BFY = Record[ nameof( BFY ) ].ToString( );
            EFY = Record[ nameof( EFY ) ].ToString( );
            FundCode = Record[ nameof( FundCode ) ].ToString( );
            FundName = Record[ nameof( FundName ) ].ToString( );
            TreasuryAccountCode = Record[ nameof( TreasuryAccountCode ) ].ToString( );
            RpioCode = Record[ nameof( RpioCode ) ].ToString( );
            RpioName = Record[ nameof( RpioName ) ].ToString( );
            Amount = double.Parse( Record[ nameof( Amount ) ].ToString( ) ?? "0" );
            OpenCommitments = double.Parse( Record[ nameof( OpenCommitments ) ].ToString( ) ?? "0" );
            Obligations = double.Parse( Record[ nameof( Obligations ) ].ToString( ) ?? "0" );
            Available = double.Parse( Record[ nameof( Available ) ].ToString( ) ?? "0" );
            TreasuryAccountCode = Record[ nameof( TreasuryAccountCode ) ].ToString( );
            TreasuryAccountName = Record[ nameof( TreasuryAccountName ) ].ToString( );
            BudgetAccountCode = Record[ nameof( BudgetAccountCode ) ].ToString( );
            BudgetAccountName = Record[ nameof( BudgetAccountName ) ].ToString( );
        }

        /// <inheritdoc/>
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.AnnualReimbursableCarryoverEstimate"/>
        /// class.
        /// </summary>
        /// <param name="builder"> The builder. </param>
        public AnnualReimbursableCarryoverEstimates( IDataModel builder )
            : this( )
        {
            Record = builder.Record;
            Data = Record.ToDictionary( );
            ID = int.Parse( Record[ "AnnualReimbursableEstimatesId" ].ToString( ) ?? "0" );
            BFY = Record[ nameof( BFY ) ].ToString( );
            EFY = Record[ nameof( EFY ) ].ToString( );
            FundCode = Record[ nameof( FundCode ) ].ToString( );
            FundName = Record[ nameof( FundName ) ].ToString( );
            TreasuryAccountCode = Record[ nameof( TreasuryAccountCode ) ].ToString( );
            RpioCode = Record[ nameof( RpioCode ) ].ToString( );
            RpioName = Record[ nameof( RpioName ) ].ToString( );
            Amount = double.Parse( Record[ nameof( Amount ) ].ToString( ) ?? "0" );
            OpenCommitments = double.Parse( Record[ nameof( OpenCommitments ) ].ToString( ) ?? "0" );
            Obligations = double.Parse( Record[ nameof( Obligations ) ].ToString( ) ?? "0" );
            Available = double.Parse( Record[ nameof( Available ) ].ToString( ) ?? "0" );
            TreasuryAccountCode = Record[ nameof( TreasuryAccountCode ) ].ToString( );
            TreasuryAccountName = Record[ nameof( TreasuryAccountName ) ].ToString( );
            BudgetAccountCode = Record[ nameof( BudgetAccountCode ) ].ToString( );
            BudgetAccountName = Record[ nameof( BudgetAccountName ) ].ToString( );
        }

        /// <inheritdoc/>
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.AnnualReimbursableCarryoverEstimate"/>
        /// class.
        /// </summary>
        /// <param name="dataRow"> The data row. </param>
        public AnnualReimbursableCarryoverEstimates( DataRow dataRow )
            : this( )
        {
            Record = dataRow;
            Data = dataRow.ToDictionary( );
            ID = int.Parse( Record[ "AnnualReimbursableEstimatesId" ].ToString( ) ?? "0" );
            BFY = dataRow[ nameof( BFY ) ].ToString( );
            EFY = dataRow[ nameof( EFY ) ].ToString( );
            FundCode = dataRow[ nameof( FundCode ) ].ToString( );
            FundName = dataRow[ nameof( FundName ) ].ToString( );
            TreasuryAccountCode = dataRow[ nameof( TreasuryAccountCode ) ].ToString( );
            RpioCode = dataRow[ nameof( RpioCode ) ].ToString( );
            RpioName = dataRow[ nameof( RpioName ) ].ToString( );
            Amount = double.Parse( dataRow[ nameof( Amount ) ].ToString( ) ?? "0" );
            OpenCommitments = double.Parse( dataRow[ nameof( OpenCommitments ) ].ToString( ) ?? "0" );
            Obligations = double.Parse( dataRow[ nameof( Obligations ) ].ToString( ) ?? "0" );
            Available = double.Parse( dataRow[ nameof( Available ) ].ToString( ) ?? "0" );
            TreasuryAccountCode = Record[ nameof( TreasuryAccountCode ) ].ToString( );
            TreasuryAccountName = Record[ nameof( TreasuryAccountName ) ].ToString( );
            BudgetAccountCode = Record[ nameof( BudgetAccountCode ) ].ToString( );
            BudgetAccountName = Record[ nameof( BudgetAccountName ) ].ToString( );
        }

        /// <inheritdoc/>
        /// <summary> </summary>
        /// <param name="carryover"> </param>
        public AnnualReimbursableCarryoverEstimates( AnnualReimbursableCarryoverEstimates carryover )
            : this( )
        {
            ID = carryover.ID;
            BFY = carryover.BFY;
            EFY = carryover.EFY;
            FundCode = carryover.FundCode;
            FundName = carryover.FundName;
            RpioCode = carryover.RpioCode;
            RpioName = carryover.RpioName;
            Amount = carryover.Amount;
            OpenCommitments = carryover.OpenCommitments;
            Obligations = carryover.Obligations;
            Available = carryover.Available;
            TreasuryAccountCode = carryover.TreasuryAccountCode;
            TreasuryAccountName = carryover.TreasuryAccountName;
            BudgetAccountCode = carryover.BudgetAccountCode;
            BudgetAccountName = carryover.BudgetAccountName;
        }
    }
}