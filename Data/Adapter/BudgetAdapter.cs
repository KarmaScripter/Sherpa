﻿// ******************************************************************************************
//     Assembly:                Budget Execution
//     Author:                  Terry D. Eppler
//     Created:                 03-24-2023
// 
//     Last Modified By:        Terry D. Eppler
//     Last Modified On:        05-31-2023
// ******************************************************************************************
// <copyright file="BudgetAdapter.cs" company="Terry D. Eppler">
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
//   BudgetAdapter.cs
// </summary>
// ******************************************************************************************

namespace BudgetExecution
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.Common;
    using System.Diagnostics.CodeAnalysis;

    /// <inheritdoc/>
    /// <summary> </summary>
    /// <seealso cref="T:BudgetExecution.AdapterBase"/>
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeInternal" ) ]
    [ SuppressMessage( "ReSharper", "ClassCanBeSealed.Global" ) ]
    public class BudgetAdapter : AdapterBase, IBudgetAdapter
    {
        /// <inheritdoc />
        /// <summary>
        /// Gets or sets the source.
        /// </summary>
        /// <value>
        /// The source.
        /// </value>
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

        /// <inheritdoc />
        /// <summary>
        /// Gets or sets the provider.
        /// </summary>
        /// <value>
        /// The provider.
        /// </value>
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

        /// <inheritdoc />
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

        /// <inheritdoc />
        /// <summary>
        /// Gets or sets the data connection.
        /// </summary>
        /// <value>
        /// The data connection.
        /// </value>
        public DbConnection DataConnection
        {
            get
            {
                return _dataConnection;
            }
            private protected set
            {
                _dataConnection = value;
            }
        }

        /// <inheritdoc />
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

        /// <inheritdoc />
        /// <summary>
        /// Gets or sets the commands.
        /// </summary>
        /// <value>
        /// The commands.
        /// </value>
        public IDictionary<string, DbCommand> Commands
        {
            get
            {
                return _commands;
            }
            private protected set
            {
                _commands = value;
            }
        }

        /// <inheritdoc />
        /// <summary>
        /// Gets or sets the command factory.
        /// </summary>
        /// <value> The command factory.
        /// </value>
        public IBudgetCommand CommandFactory
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

        /// <inheritdoc />
        /// <summary>
        /// Gets or sets the command text.
        /// </summary>
        /// <value>
        /// The command text.
        /// </value>
        public string CommandText
        {
            get
            {
                return _commandText;
            }
            private protected set
            {
                _commandText = value;
            }
        }

        /// <inheritdoc/>
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetAdapter"/>
        /// class.
        /// </summary>
        public BudgetAdapter( )
        {
            MissingSchemaAction = MissingSchemaAction.AddWithKey;
            MissingMappingAction = MissingMappingAction.Ignore;
            ContinueUpdateOnError = true;
            AcceptChangesDuringFill = true;
            AcceptChangesDuringUpdate = true;
            ReturnProviderSpecificTypes = true;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetAdapter" /> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="provider">The provider.</param>
        /// <param name="commandType">Type of the command.</param>
        public BudgetAdapter( Source source, Provider provider, SQL commandType = SQL.SELECTALL )
            : this( )
        {
            _source = source;
            _provider = provider;
            _sqlStatement = new SqlStatement( source, provider, commandType );
            _dataConnection = new BudgetConnection( source, provider ).Create( );
            _commandText = _sqlStatement.CommandText;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetAdapter" /> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="provider">The provider.</param>
        /// <param name="sqlText">The SQL text.</param>
        /// <param name="commandType">Type of the command.</param>
        public BudgetAdapter( Source source, Provider provider, string sqlText, SQL commandType )
            : this( )
        {
            _source = source;
            _provider = provider;
            _sqlStatement = new SqlStatement( source, provider, commandType );
            _dataConnection = new BudgetConnection( source, provider ).Create( );
            _commandText = _sqlStatement.CommandText;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the <see cref="T:BudgetExecution.BudgetAdapter" /> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="provider">The provider.</param>
        /// <param name="where">The where.</param>
        /// <param name="commandType">Type of the command.</param>
        public BudgetAdapter( Source source, Provider provider, IDictionary<string, object> where,
            SQL commandType = SQL.SELECTALL )
            : this( )
        {
            _source = source;
            _provider = provider;
            _sqlStatement = new SqlStatement( source, provider, where, commandType );
            _dataConnection = new BudgetConnection( source, provider ).Create( );
            _commandText = _sqlStatement.CommandText;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the <see cref="T:BudgetExecution.BudgetAdapter" /> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="provider">The provider.</param>
        /// <param name="update">The update.</param>
        /// <param name="where">The where.</param>
        /// <param name="commandType">Type of the command.</param>
        public BudgetAdapter( Source source, Provider provider, IDictionary<string, object> update,
            IDictionary<string, object> where, SQL commandType = SQL.UPDATE )
            : this( )
        {
            _source = source;
            _provider = provider;
            _sqlStatement = new SqlStatement( source, provider, update, where, commandType );
            _dataConnection = new BudgetConnection( source, provider ).Create( );
            _commandText = _sqlStatement.CommandText;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetAdapter" /> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="provider">The provider.</param>
        /// <param name="columns">The columns.</param>
        /// <param name="where">The where.</param>
        /// <param name="commandType">Type of the command.</param>
        public BudgetAdapter( Source source, Provider provider, IEnumerable<string> columns,
            IDictionary<string, object> where, SQL commandType = SQL.SELECT )
            : this( )
        {
            _source = source;
            _provider = provider;
            _sqlStatement = new SqlStatement( source, provider, columns, where, commandType );
            _dataConnection = new BudgetConnection( source, provider ).Create( );
            _commandText = _sqlStatement.CommandText;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetAdapter" /> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="provider">The provider.</param>
        /// <param name="fields">The fields.</param>
        /// <param name="numerics">The numerics.</param>
        /// <param name="having">The having.</param>
        /// <param name="commandType">Type of the command.</param>
        public BudgetAdapter( Source source, Provider provider, IEnumerable<string> fields,
            IEnumerable<string> numerics, IDictionary<string, object> having,
            SQL commandType = SQL.SELECT )
            : this( )
        {
            _source = source;
            _provider = provider;
            _sqlStatement = new SqlStatement( source, provider, fields, numerics, having,
                commandType );

            _dataConnection = new BudgetConnection( source, provider ).Create( );
            _commandText = _sqlStatement.CommandText;
        }

        /// <inheritdoc/>
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetAdapter"/>
        /// class.
        /// </summary>
        /// <param name="sqlStatement"> The SQL statement. </param>
        public BudgetAdapter( ISqlStatement sqlStatement )
            : this( )
        {
            _source = sqlStatement.Source;
            _provider = sqlStatement.Provider;
            _sqlStatement = sqlStatement;
            _dataConnection = new BudgetConnection( _source, _provider ).Create( );
            _commandText = sqlStatement.CommandText;
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:BudgetExecution.BudgetAdapter" /> class.
        /// </summary>
        /// <param name="command">The command.</param>
        public BudgetAdapter( IBudgetCommand command )
            : this( )
        {
            _source = command.Source;
            _provider = command.Provider;
            _sqlStatement = command.SqlStatement;
            _dataConnection = command.Connection;
            _commandText = command.SqlStatement.CommandText;
        }
    }
}