﻿// <copyright file = "TravelActivity.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>

namespace BudgetExecution
{
    using System;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// 
    /// </summary>
    /// <seealso cref = "Obligation"/>
    [ SuppressMessage( "ReSharper", "UnusedType.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    public class TravelActivity : TravelData
    {
        /// <summary>
        /// Gets or sets the source.
        /// </summary>
        /// <value>
        /// The source.
        /// </value>
        public override Source Source { get; set; } =  Source.TravelActivity;

        /// <summary>
        /// Initializes a new instance of the <see cref = "TravelActivity"/> class.
        /// </summary>
        /// <inheritdoc/>
        public TravelActivity()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "TravelActivity"/> class.
        /// </summary>
        /// <param name = "query" >
        /// </param>
        public TravelActivity( IQuery query )
            : base( query )
        {
            Record = new DataBuilder( query )?.GetRecord();
            ID = new Key( Record, PrimaryKey.TravelActivityId );
            FocCode = new Element( Record, Field.FocCode );
            FocName = new Element( Record, Field.FocName );
            DCN = new Element( Record, Field.DocumentControlNumber );
            FirstName = new Element( Record, Field.FirstName );
            LastName = new Element( Record, Field.LastName );
            StartDate = new Time( Record, EventDate.StartDate );
            EndDate = new Time( Record, EventDate.EndDate );
            Obligations = new Amount( Record, Numeric.Obligations );
            ULO = new Amount( Record, Numeric.ULO );
            Expenditures = new Amount( Record, Numeric.Expenditures );
            Data = Record?.ToDictionary();
            Type = ExpenseType.Obligation;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "TravelActivity"/> class.
        /// </summary>
        /// <param name = "builder" >
        /// The builder.
        /// </param>
        public TravelActivity( IBuilder builder )
            : base( builder )
        {
            Record = builder.GetRecord();
            ID = new Key( Record, PrimaryKey.TravelActivityId );
            FocCode = new Element( Record, Field.FocCode );
            FocName = new Element( Record, Field.FocName );
            DCN = new Element( Record, Field.DocumentControlNumber );
            FirstName = new Element( Record, Field.FirstName );
            LastName = new Element( Record, Field.LastName );
            StartDate = new Time( Record, EventDate.StartDate );
            EndDate = new Time( Record, EventDate.EndDate );
            Obligations = new Amount( Record, Numeric.Obligations );
            ULO = new Amount( Record, Numeric.ULO );
            Expenditures = new Amount( Record, Numeric.Expenditures );
            Data = Record?.ToDictionary();
            Type = ExpenseType.Obligation;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "TravelActivity"/> class.
        /// </summary>
        /// <param name = "dataRow" >
        /// The dr.
        /// </param>
        public TravelActivity( DataRow dataRow )
        {
            Record = dataRow;
            ID = new Key( Record, PrimaryKey.TravelActivityId );
            FocCode = new Element( Record, Field.FocCode );
            FocName = new Element( Record, Field.FocName );
            DCN = new Element( Record, Field.DocumentControlNumber );
            FirstName = new Element( Record, Field.FirstName );
            LastName = new Element( Record, Field.LastName );
            StartDate = new Time( Record, EventDate.StartDate );
            EndDate = new Time( Record, EventDate.EndDate );
            Obligations = new Amount( Record, Numeric.Obligations );
            ULO = new Amount( Record, Numeric.ULO );
            Expenditures = new Amount( Record, Numeric.Expenditures );
            Data = Record?.ToDictionary();
            Type = ExpenseType.Obligation;
        }
    }
}
