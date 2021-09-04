﻿// <copyright file = "Employee.cs" company = "Terry D. Eppler">
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
    /// <seealso cref = "IEmployee"/>
    /// <seealso cref = "ISource"/>
    [ SuppressMessage( "ReSharper", "PublicConstructorInAbstractClass" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeInternal" ) ]
    [ SuppressMessage( "ReSharper", "AutoPropertyCanBeMadeGetOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "ConvertToConstant.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeProtected.Global" ) ]
    public class Employee : EmployeeBase, IEmployee, ISource
    {
        /// <summary>
        /// Gets or sets the source.
        /// </summary>
        /// <value>
        /// The source.
        /// </value>
        private const Source _source = Source.Employees;
        
        /// <summary>
        /// Initializes a new instance of the <see cref = "Employee"/> class.
        /// </summary>
        public Employee()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "Employee"/> class.
        /// </summary>
        /// <param name = "query" >
        /// The query.
        /// </param>
        public Employee( IQuery query )
        {
            _record = new DataBuilder( query )?.GetRecord();
            _id = new Key( _record, PrimaryKey.EmployeeId );
            _section = new Element( _record, Field.Section );
            _firstName = new Element( _record, Field.FirstName );
            _lastName = new Element( _record, Field.LastName );
            _employeeNumber = new Element( _record, Field.EmployeeNumber );
            _office = new Element( _record, Field.Office );
            _phoneNumber = new Element( _record, Field.PhoneNumber );
            _cellNumber = new Element( _record, Field.CellNumber );
            _email = new Element( _record, Field.Email );
            _status = new Element( _record, Field.Status );
            _args = _record?.ToDictionary();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "Employee"/> class.
        /// </summary>
        /// <param name = "builder" >
        /// The builder.
        /// </param>
        public Employee( IBuilder builder )
        {
            _record = builder?.GetRecord();
            _id = new Key( _record, PrimaryKey.EmployeeId );
            _section = new Element( _record, Field.Section );
            _firstName = new Element( _record, Field.FirstName );
            _lastName = new Element( _record, Field.LastName );
            _employeeNumber = new Element( _record, Field.EmployeeNumber );
            _office = new Element( _record, Field.Office );
            _phoneNumber = new Element( _record, Field.PhoneNumber );
            _cellNumber = new Element( _record, Field.CellNumber );
            _email = new Element( _record, Field.Email );
            _status = new Element( _record, Field.Status );
            _args = _record?.ToDictionary();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "Employee"/> class.
        /// </summary>
        /// <param name = "dataRow" >
        /// The dataRow.
        /// </param>
        public Employee( DataRow dataRow )
        {
            _record = dataRow;
            _id = new Key( _record, PrimaryKey.EmployeeId );
            _section = new Element( _record, Field.Section );
            _firstName = new Element( _record, Field.FirstName );
            _lastName = new Element( _record, Field.LastName );
            _employeeNumber = new Element( _record, Field.EmployeeNumber );
            _office = new Element( _record, Field.Office );
            _phoneNumber = new Element( _record, Field.PhoneNumber );
            _cellNumber = new Element( _record, Field.CellNumber );
            _email = new Element( _record, Field.Email );
            _status = new Element( _record, Field.Status );
            _args = _record?.ToDictionary();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref = "Employee"/> class.
        /// </summary>
        /// <param name = "epaNumber" >
        /// The epaNumber.
        /// </param>
        public Employee( string epaNumber )
        {
            _record = new DataBuilder( _source, SetArgs( epaNumber ) )?.GetRecord();
            _id = new Key( _record, PrimaryKey.EmployeeId );
            _section = new Element( _record, Field.Section );
            _firstName = new Element( _record, Field.FirstName );
            _lastName = new Element( _record, Field.LastName );
            _employeeNumber = new Element( _record, Field.EmployeeNumber );
            _office = new Element( _record, Field.Office );
            _phoneNumber = new Element( _record, Field.PhoneNumber );
            _cellNumber = new Element( _record, Field.CellNumber );
            _email = new Element( _record, Field.Email );
            _status = new Element( _record, Field.Status );
            _args = _record?.ToDictionary();
        }
        
        /// <summary>
        /// Sets the arguments.
        /// </summary>
        /// <param name = "code" >
        /// The code.
        /// </param>
        /// <returns>
        /// </returns>
        public IDictionary<string, object> SetArgs( string code )
        {
            if( Verify.Input( code ) )
            {
                try
                {
                    return new Dictionary<string, object>
                    {
                        [ $"{Field.EmployeeNumber}" ] = code
                    };
                }
                catch( Exception ex )
                {
                    Fail( ex );
                    return default( IDictionary<string, object> );
                }
            }

            return default( IDictionary<string, object> );
        }

        /// <summary>
        /// Gets the first name.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetFirstName()
        {
            try
            {
                return Verify.Element( _firstName )
                    ? _firstName
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the last name.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetLastName()
        {
            try
            {
                return Verify.Element( _lastName )
                    ? _lastName
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the employee identifier.
        /// </summary>
        /// <returns>
        /// </returns>
        public IKey GetId()
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
        /// Gets the section.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetSection()
        {
            try
            {
                return Verify.Element( _section )
                    ? _section
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the employee number.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetEmployeeNumber()
        {
            try
            {
                return Verify.Element( _employeeNumber )
                    ? _employeeNumber
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the office.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetOffice()
        {
            try
            {
                return Verify.Element( _office )
                    ? _office
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the phone.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetPhoneNumber()
        {
            try
            {
                return Verify.Element( _phoneNumber )
                    ? _phoneNumber
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the cell.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetCellNumber()
        {
            try
            {
                return Verify.Element( _cellNumber )
                    ? _cellNumber
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the email.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetEmail()
        {
            try
            {
                return Verify.Element( _email )
                    ? _email
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Gets the employement status.
        /// </summary>
        /// <returns>
        /// </returns>
        public IElement GetEmployementStatus()
        {
            try
            {
                return Verify.Element( _status )
                    ? _status
                    : Element.Default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Element.Default;
            }
        }

        /// <summary>
        /// Converts to dictionary.
        /// </summary>
        /// <returns>
        /// </returns>
        public IDictionary<string, object> ToDictionary()
        {
            try
            {
                return Verify.Map( _args )
                    ? _args
                    : default( IDictionary<string, object> );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( IDictionary<string, object> );
            }
        }

        /// <summary>
        /// Gets the source.
        /// </summary>
        /// <returns>
        /// </returns>
        public Source GetSource()
        {
            try
            {
                return Validate.Source( _source )
                    ? _source
                    : Source.NS;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return Source.NS;
            }
        }
    }
}
