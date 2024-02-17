﻿// ************************************************************************************************
//     Assembly:                Budget Execution
//     Author:                  Terry D. Eppler
//     Created:              19-10-2023
// 
//     Last Modified By:        Terry D. Eppler
//     Last Modified On:        13-02-2024
// ************************************************************************************************
// <copyright file="CommandBase.cs.cs" company="Terry D. Eppler">
//    This is a Federal Budget, Finance, and Accounting application for analysts in the
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
//    You can contact me at: terryeppler@gmail.com or eppler.terry@epa.gov
// </copyright>
// <summary>
//   Eppler, Terry.cs
// </summary>
// ************************************************************************************************

namespace BudgetExecution
{
    using System;
    using System.Data.Common;
    using System.Data.OleDb;
    using System.Data.SqlClient;
    using System.Data.SQLite;
    using System.Data.SqlServerCe;
    using System.Diagnostics.CodeAnalysis;

    /// <inheritdoc/>
    /// <summary> </summary>
    /// <seealso cref="T:BudgetExecution.ISource"/>
    /// <seealso cref="T:BudgetExecution.IProvider"/>
    [ SuppressMessage( "ReSharper", "PropertyCanBeMadeInitOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeProtected.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "AutoPropertyCanBeMadeGetOnly.Global" ) ]
    [ SuppressMessage( "ReSharper", "ConvertToAutoProperty" ) ]
    [ SuppressMessage( "ReSharper", "InconsistentNaming" ) ]
    public abstract class CommandBase : ISource, IProvider
    {
        /// <summary>
        /// The source
        /// </summary>
        /// <value>
        ///  Source
        /// </value>
        private protected Source _source;

        /// <summary>
        /// The provider
        /// </summary>
        /// <value>
        ///  Provider
        /// </value>
        private protected Provider _provider;

        /// <summary>
        ///  
        /// </summary>
        /// <value>
        ///  
        /// </value>
        private protected DbCommand _command;

        /// <summary>
        ///  
        /// </summary>
        /// <value>
        ///  
        /// </value>
        private protected DbConnection _connection;
        
        /// <summary>
        /// 
        /// </summary>
        /// <value>
        ///  
        /// </value>
        private protected SQL _commandType;

        /// <summary>
        ///  
        /// </summary>
        /// <value>
        /// 
        /// </value>
        private protected ISqlStatement _sqlStatement;
        
        /// <summary>
        /// Gets or sets the command.
        /// </summary>
        /// <value>
        /// command.
        /// </value>
        public DbCommand Command
        {
            get
            {
                return _command;
            }
            private protected set
            {
                _command = value;
            }
        }

        /// <summary>
        /// Gets or sets the connection factory.
        /// </summary>
        /// <value>
        /// The connection factory.
        /// </value>
        public DbConnection Connection
        {
            get
            {
                return _connection;
            }
            private protected set
            {
                _connection = value;
            }
        }

        /// <summary>
        /// Gets or sets the type of the command.
        /// </summary>
        /// <value>
        /// The type of the command.
        /// </value>
        public SQL CommandType
        {
            get
            {
                return _commandType;
            }
            private protected set
            {
                _commandType = value;
            }
        }

        /// <summary>
        /// Gets or sets the SQL statement.
        /// </summary>
        /// <value>
        /// The SQL statement.
        /// </value>
        public ISqlStatement SqlStatement
        {
            get
            {
                return _sqlStatement;
            }
            private protected set
            {
                _sqlStatement = value;
            }
        }

        /// <inheritdoc/>
        /// <summary>
        /// </summary>
        public Provider Provider
        {
            get
            {
                return _provider;
            }
            private protected set
            {
                _provider = value;
            }
        }

        /// <inheritdoc/>
        /// <summary>
        /// Gets the source.
        /// </summary>
        public Source Source
        {
            get
            {
                return _source;
            }
            private protected set
            {
                _source = value;
            }
        }

        /// <summary>
        /// Gets the sqlite command.
        /// </summary>
        /// <returns> </returns>
        private protected DbCommand GetSQLiteCommand( )
        {
            if( ( _sqlStatement != null )
               && ( _connection != null ) )
            {
                try
                {
                    switch( _sqlStatement?.CommandType )
                    {
                        case SQL.SELECTALL:
                        case SQL.SELECT:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return !string.IsNullOrEmpty( _sql )
                                ? new SQLiteCommand( _sql, _connection as SQLiteConnection )
                                : default( SQLiteCommand );
                        }
                        case SQL.INSERT:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SQLiteCommand( _sql, _connection as SQLiteConnection );
                        }
                        case SQL.UPDATE:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SQLiteCommand( _sql, _connection as SQLiteConnection );
                        }
                        case SQL.DELETE:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SQLiteCommand( _sql, _connection as SQLiteConnection );
                        }
                        default:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SQLiteCommand( _sql, _connection as SQLiteConnection );
                        }
                    }
                }
                catch( Exception _ex )
                {
                    Fail( _ex );
                    return default( DbCommand );
                }
            }

            return default( DbCommand );
        }

        /// <summary>
        /// Gets the SQL ce command.
        /// </summary>
        /// <returns>
        /// </returns>
        private protected DbCommand GetSqlCeCommand( )
        {
            if( ( _sqlStatement != null )
               && ( _connection != null ) )
            {
                try
                {
                    switch( _sqlStatement?.CommandType )
                    {
                        case SQL.SELECTALL:
                        case SQL.SELECT:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return !string.IsNullOrEmpty( _sql )
                                ? new SqlCeCommand( _sql, _connection as SqlCeConnection )
                                : default( DbCommand );
                        }
                        case SQL.INSERT:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return !string.IsNullOrEmpty( _sql )
                                ? new SqlCeCommand( _sql, _connection as SqlCeConnection )
                                : default( DbCommand );
                        }
                        case SQL.UPDATE:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return !string.IsNullOrEmpty( _sql )
                                ? new SqlCeCommand( _sql, _connection as SqlCeConnection )
                                : default( DbCommand );
                        }
                        case SQL.DELETE:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return !string.IsNullOrEmpty( _sql )
                                ? new SqlCeCommand( _sql, _connection as SqlCeConnection )
                                : default( DbCommand );
                        }
                        default:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return !string.IsNullOrEmpty( _sql )
                                ? new SqlCeCommand( _sql, _connection as SqlCeConnection )
                                : default( DbCommand );
                        }
                    }
                }
                catch( Exception _ex )
                {
                    Fail( _ex );
                    return default( DbCommand );
                }
            }

            return default( DbCommand );
        }

        /// <summary>
        /// Gets the SQL command.
        /// </summary>
        /// <returns>
        /// </returns>
        private protected DbCommand GetSqlCommand( )
        {
            if( ( _sqlStatement != null )
               && ( _connection != null ) )
            {
                try
                {
                    switch( _sqlStatement?.CommandType )
                    {
                        case SQL.SELECTALL:
                        case SQL.SELECT:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SqlCommand( _sql, _connection as SqlConnection );
                        }
                        case SQL.INSERT:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SqlCommand( _sql, _connection as SqlConnection );
                        }
                        case SQL.UPDATE:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SqlCommand( _sql, _connection as SqlConnection );
                        }
                        case SQL.DELETE:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SqlCommand( _sql, _connection as SqlConnection );
                        }
                        default:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new SqlCommand( _sql, _connection as SqlConnection );
                        }
                    }
                }
                catch( Exception _ex )
                {
                    Fail( _ex );
                    return default( DbCommand );
                }
            }

            return default( DbCommand );
        }

        /// <summary> Gets the OLE database command. </summary>
        /// <returns> </returns>
        private protected DbCommand GetOleDbCommand( )
        {
            if( ( _sqlStatement != null )
               && Enum.IsDefined( typeof( SQL ), _sqlStatement.CommandType )
               && ( _connection != null ) )
            {
                try
                {
                    switch( _sqlStatement?.CommandType )
                    {
                        case SQL.SELECTALL:
                        case SQL.SELECT:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new OleDbCommand( _sql, _connection as OleDbConnection );
                        }
                        case SQL.INSERT:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new OleDbCommand( _sql, _connection as OleDbConnection );
                        }
                        case SQL.UPDATE:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new OleDbCommand( _sql, _connection as OleDbConnection );
                        }
                        case SQL.DELETE:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new OleDbCommand( _sql, _connection as OleDbConnection );
                        }
                        default:
                        {
                            var _sql = _sqlStatement?.GetCommandText( );
                            return new OleDbCommand( _sql, _connection as OleDbConnection );
                        }
                    }
                }
                catch( Exception _ex )
                {
                    Fail( _ex );
                    return default( DbCommand );
                }
            }

            return default( DbCommand );
        }

        /// <summary> Fails the specified ex. </summary>
        /// <param name="ex"> The ex. </param>
        protected static void Fail( Exception ex )
        {
            using var _error = new ErrorDialog( ex );
            _error?.SetText( );
            _error?.ShowDialog( );
        }
    }
}