﻿// ******************************************************************************************
//     Assembly:                Budget Execution
//     Author:                  Terry D. Eppler
//     Created:                 03-24-2023
// 
//     Last Modified By:        Terry D. Eppler
//     Last Modified On:        05-31-2023
// ******************************************************************************************
// <copyright file="AdapterBase.cs" company="Terry D. Eppler">
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
//   AdapterBase.cs
// </summary>
// ******************************************************************************************

namespace BudgetExecution
{
    using System;
    using System.Collections.Generic;
    using System.Data.Common;
    using System.Data.OleDb;
    using System.Data.SqlClient;
    using System.Data.SQLite;
    using System.Data.SqlServerCe;
    using System.Diagnostics.CodeAnalysis;

    /// <inheritdoc/>
    /// <summary> </summary>
    /// <seealso cref="T:System.Data.Common.DbDataAdapter"/>
    /// <seealso cref="T:BudgetExecution.ISource"/>
    /// <seealso cref="T:BudgetExecution.IProvider"/>
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "ArrangeDefaultValueWhenTypeNotEvident" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeProtected.Global" ) ]
    [ SuppressMessage( "ReSharper", "VirtualMemberNeverOverridden.Global" ) ]
    [ SuppressMessage( "ReSharper", "AutoPropertyCanBeMadeGetOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "MergeConditionalExpression" ) ]
    [ SuppressMessage( "ReSharper", "PropertyCanBeMadeInitOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "InconsistentNaming" ) ]
    [ SuppressMessage( "ReSharper", "FieldCanBeMadeReadOnly.Global" ) ]
    public class AdapterBase : DbDataAdapter
    {
        /// <summary>
        /// The source
        /// </summary>
        private protected Source _source;

        /// <summary>
        /// The provider
        /// </summary>
        private protected Provider _provider;

        /// <summary>
        /// The command type
        /// </summary>
        private protected Command _commandType;

        /// <summary>
        /// The data connection
        /// </summary>
        private protected DbConnection _dataConnection;

        /// <summary>
        /// The SQL statement
        /// </summary>
        private protected ISqlStatement _sqlStatement;

        /// <summary>
        /// The criteria
        /// </summary>
        private protected IDictionary<string, object> _criteria;

        /// <summary>
        /// The commands
        /// </summary>
        private protected IDictionary<string, DbCommand> _commands;

        /// <summary>
        /// The command factory
        /// </summary>
        private protected IBudgetCommand _command;

        /// <summary>
        /// The command text
        /// </summary>
        private protected string _commandText;

        /// <summary>
        /// Gets the sql lite adapter.
        /// </summary>
        /// <returns></returns>
        private protected SQLiteDataAdapter CreateSQLiteAdapter( )
        {
            try
            {
                var _adapter =
                    new SQLiteDataAdapter( _commandText, _dataConnection as SQLiteConnection );

                _adapter.ContinueUpdateOnError = true;
                _adapter.AcceptChangesDuringFill = true;
                _adapter.AcceptChangesDuringUpdate = true;
                _adapter.ReturnProviderSpecificTypes = true;
                if( _commandText.StartsWith( @"SELECT *" )
                   | _commandText.StartsWith( "SELECT ALL" ) )
                {
                    var _builder = new SQLiteCommandBuilder( _adapter );
                    _adapter.InsertCommand = _builder.GetInsertCommand( );
                    _adapter.UpdateCommand = _builder.GetUpdateCommand( );
                    _adapter.DeleteCommand = _builder.GetDeleteCommand( );
                    return _adapter;
                }
                else
                {
                    return _adapter;
                }
            }
            catch( Exception _ex )
            {
                Fail( _ex );
                return default( SQLiteDataAdapter );
            }
        }

        /// <summary>
        /// Gets the SQL adapter.
        /// </summary>
        /// <returns></returns>
        private protected SqlDataAdapter CreateSqlDataAdapter( )
        {
            try
            {
                var _adapter = 
                    new SqlDataAdapter( _commandText, _dataConnection as SqlConnection );

                _adapter.ContinueUpdateOnError = true;
                _adapter.AcceptChangesDuringFill = true;
                _adapter.AcceptChangesDuringUpdate = true;
                _adapter.ReturnProviderSpecificTypes = true;
                if( _commandText.StartsWith( @"SELECT *" )
                   | _commandText.StartsWith( "SELECT ALL" ) )
                {
                    var _builder = new SqlCommandBuilder( _adapter );
                    _adapter.InsertCommand = _builder.GetInsertCommand( );
                    _adapter.UpdateCommand = _builder.GetUpdateCommand( );
                    _adapter.DeleteCommand = _builder.GetDeleteCommand( );
                    return _adapter;
                }
                else
                {
                    return _adapter;
                }
            }
            catch( Exception _ex )
            {
                Fail( _ex );
                return default( SqlDataAdapter );
            }
        }

        /// <summary>
        /// Gets the OLE database adapter.
        /// </summary>
        /// <returns>
        /// </returns>
        private protected OleDbDataAdapter CreateOleDbAdapter( )
        {
            try
            {
                var _adapter = 
                    new OleDbDataAdapter( _commandText, _dataConnection as OleDbConnection );

                _adapter.ContinueUpdateOnError = true;
                _adapter.AcceptChangesDuringFill = true;
                _adapter.AcceptChangesDuringUpdate = true;
                _adapter.ReturnProviderSpecificTypes = true;
                if( _commandText.StartsWith( @"SELECT *" ) 
                   | _commandText.StartsWith( "SELECT ALL" ) )
                {
                    var _builder = new OleDbCommandBuilder( _adapter );
                    _adapter.InsertCommand = _builder.GetInsertCommand( );
                    _adapter.UpdateCommand = _builder.GetUpdateCommand( );
                    _adapter.DeleteCommand = _builder.GetDeleteCommand( );
                    return ( _adapter != null )
                        ? _adapter
                        : default( OleDbDataAdapter );
                }
                else
                {
                    return _adapter;
                }
            }
            catch( Exception _ex )
            {
                Fail( _ex );
                return default( OleDbDataAdapter );
            }
        }

        /// <summary>
        /// Gets the SQL ce adapter.
        /// </summary>
        /// <returns> </returns>
        private protected DbDataAdapter CreateSqlCompactAdapter( )
        {
            try
            {
                var _adapter = 
                    new SqlCeDataAdapter( _commandText, _dataConnection as SqlCeConnection );

                _adapter.ContinueUpdateOnError = true;
                _adapter.AcceptChangesDuringFill = true;
                _adapter.AcceptChangesDuringUpdate = true;
                _adapter.ReturnProviderSpecificTypes = true;
                if( _commandText.StartsWith( @"SELECT *" )
                   | _commandText.StartsWith( "SELECT ALL" ) )
                {
                    var _builder = new SqlCeCommandBuilder( _adapter );
                    _adapter.InsertCommand = _builder.GetInsertCommand( );
                    _adapter.UpdateCommand = _builder.GetUpdateCommand( );
                    _adapter.DeleteCommand = _builder.GetDeleteCommand( );
                    return _adapter;
                }
                else
                {
                    return _adapter;
                }
            }
            catch( Exception _ex )
            {
                Fail( _ex );
                return default( SqlCeDataAdapter );
            }
        }

        /// <summary>
        /// Fails the specified ex.
        /// </summary>
        /// <param name="ex">
        /// The ex.
        /// </param>
        private protected void Fail( Exception ex )
        {
            using var _error = new ErrorDialog( ex );
            _error?.SetText( );
            _error?.ShowDialog( );
        }
    }
}