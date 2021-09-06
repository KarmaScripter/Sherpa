﻿// <copyright file = "Vendor.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>

namespace BudgetExecution
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// 
    /// </summary>
    /// <seealso cref = "Obligation"/>
    /// <seealso cref = "IVendor"/>
    /// <seealso cref = "Obligation"/>
    /// <seealso cref = "IVendor"/>
    /// <seealso cref = "IDataBuilder"/>
    [ SuppressMessage( "ReSharper", "AssignNullToNotNullAttribute" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    public class Vendor : VendorData, IVendor
    {
        /// <summary>
        /// Gets or sets the source.
        /// </summary>
        /// <value>
        /// The source.
        /// </value>
        private protected new Source _source = Source.Vendors;

        /// <summary>
        /// Initializes a new instance of the <see cref = "Vendor"/> class.
        /// </summary>
        /// <inheritdoc/>
        public Vendor()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "Vendor"/> class.
        /// </summary>
        /// <param name = "query" >
        /// </param>
        public Vendor( IQuery query )
            : base( query )
        {
            _records = new DataBuilder( query )?.GetRecord();
            _id = new Key( _records, PrimaryKey.VendorId );
            _code = new Element( _records, Field.Code );
            _name = new Element( _records, Field.Name );
            _dunsNumber = new Element( _records, Field.DunsNumber );
            _programProjectCode = new Element( _records, Field.ProgramProjectCode );
            _focCode = new Element( _records, Field.FocCode );
            _focName = new Element( _records, Field.FocName );
            _documentType = new Element( _records, Field.DocumentType );
            _documentNumber = new Element( _records, Field.DocumentNumber );
            _startDate = new Time( _records, EventDate.StartDate );
            _endDate = new Time( _records, EventDate.EndDate );
            _closedDate = new Time( _records, EventDate.ClosedDate );
            _amount = new Amount( _records, Numeric.Amount );
            _expended = new Amount( _records, Numeric.Expended );
            _ulo = new Amount( _records, Numeric.ULO );
            _data = _records?.ToDictionary();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "Vendor"/> class.
        /// </summary>
        /// <param name = "builder" >
        /// The builder.
        /// </param>
        public Vendor( IBuilder builder )
            : base( builder )
        {
            _records = builder?.GetRecord();
            _id = new Key( _records, PrimaryKey.VendorId );
            _code = new Element( _records, Field.Code );
            _name = new Element( _records, Field.Name );
            _dunsNumber = new Element( _records, Field.DunsNumber );
            _programProjectCode = new Element( _records, Field.ProgramProjectCode );
            _focCode = new Element( _records, Field.FocCode );
            _focName = new Element( _records, Field.FocName );
            _documentType = new Element( _records, Field.DocumentType );
            _documentNumber = new Element( _records, Field.DocumentNumber );
            _startDate = new Time( _records, EventDate.StartDate );
            _endDate = new Time( _records, EventDate.EndDate );
            _closedDate = new Time( _records, EventDate.ClosedDate );
            _amount = new Amount( _records, Numeric.Amount );
            _expended = new Amount( _records, Numeric.Expended );
            _ulo = new Amount( _records, Numeric.ULO );
            _data = _records?.ToDictionary();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "Vendor"/> class.
        /// </summary>
        /// <param name = "datarow" >
        /// The dr.
        /// </param>
        public Vendor( DataRow datarow )
            : this()
        {
            _records = datarow;
            _id = new Key( _records, PrimaryKey.VendorId );
            _code = new Element( _records, Field.Code );
            _name = new Element( _records, Field.Name );
            _dunsNumber = new Element( _records, Field.DunsNumber );
            _programProjectCode = new Element( _records, Field.ProgramProjectCode );
            _focCode = new Element( _records, Field.FocCode );
            _focName = new Element( _records, Field.FocName );
            _documentType = new Element( _records, Field.DocumentType );
            _documentNumber = new Element( _records, Field.DocumentNumber );
            _startDate = new Time( _records, EventDate.StartDate );
            _endDate = new Time( _records, EventDate.EndDate );
            _closedDate = new Time( _records, EventDate.ClosedDate );
            _amount = new Amount( _records, Numeric.Amount );
            _expended = new Amount( _records, Numeric.Expended );
            _ulo = new Amount( _records, Numeric.ULO );
            _data = _records?.ToDictionary();
        }
        
        /// <summary>
        /// Gets the vendor identifier.
        /// </summary>
        /// <returns>
        /// </returns>
        public override IKey GetId()
        {
            try
            {
                return Verify.Key( _id )
                    ? _id
                    : Key.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Key.Default;
            }
        }

        /// <summary>
        /// Gets the vendor code.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetCode()
        {
            try
            {
                return Verify.Element( _code )
                    ? _code
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the name of the vendor.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetName()
        {
            try
            {
                return Verify.Element( _name )
                    ? _name
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Converts to string.
        /// </summary>
        /// <returns>
        /// A <see cref = "string"/> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            try
            {
                return Verify.Element( _name )
                    ? _name?.GetValue()
                    : string.Empty;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the duns number.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetDunsNumber()
        {
            try
            {
                return Verify.Element( _dunsNumber )
                    ? _dunsNumber
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the document number.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetDocumentNumber()
        {
            try
            {
                return Verify.Element( _documentNumber )
                    ? _documentNumber
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the start date.
        /// </summary>
        /// <returns>
        /// </returns>
        public ITime GetStartDate()
        {
            try
            {
                return Verify.Time( _startDate )
                    ? _startDate
                    : Time.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Time.Default;
            }
        }

        /// <summary>
        /// Gets the end date.
        /// </summary>
        /// <returns>
        /// </returns>
        public ITime GetEndDate()
        {
            try
            {
                return Verify.Time( _endDate )
                    ? _endDate
                    : Time.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Time.Default;
            }
        }

        /// <summary>
        /// Gets the closed date.
        /// </summary>
        /// <returns>
        /// </returns>
        public ITime GetClosedDate()
        {
            try
            {
                return Verify.Time( _closedDate )
                    ? _closedDate
                    : Time.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Time.Default;
            }
        }

        /// <summary>
        /// Gets the expended.
        /// </summary>
        /// <returns>
        /// </returns>
        public IAmount GetExpended()
        {
            try
            {
                return Verify.Amount( _expended )
                    ? _expended
                    : default( IAmount );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( IAmount );
            }
        }

        /// <summary>
        /// Gets the unliquidated obligation.
        /// </summary>
        /// <returns>
        /// </returns>
        public IAmount GetUnliquidatedObligation()
        {
            try
            {
                return Verify.Amount( _ulo )
                    ? _ulo
                    : default( IAmount );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( IAmount );
            }
        }

        /// <summary>
        /// Converts to dictionary.
        /// </summary>
        /// <returns>
        /// </returns>
        public override IDictionary<string, object> ToDictionary()
        {
            try
            {
                return Verify.Map( _data )
                    ? _data
                    : default( IDictionary<string, object> );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( IDictionary<string, object> );
            }
        }
    }
}
