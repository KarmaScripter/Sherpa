﻿// <copyright file = "BindingData.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>

namespace BudgetExecution
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Linq;
    using System.Threading;
    using System.Windows.Forms;

    /// <summary> </summary>
    /// <seealso cref = "BudgetExecution.BindingBase"/>
    public class BindingData : BindingBase
    {
        /// <summary>
        /// Gets the data set.
        /// </summary>
        /// <value>
        /// The data set.
        /// </value>
        public override DataSet DataSet { get; set; }

        /// <summary>
        /// Gets the data table.
        /// </summary>
        /// <value>
        /// The data table.
        /// </value>
        public override DataTable DataTable { get; set; }

        /// <summary>
        /// Gets the source.
        /// </summary>
        /// <value>
        /// The source.
        /// </value>
        public override Source Source { get; set; }

        /// <summary>
        /// Gets or sets the current.
        /// </summary>
        /// <value>
        /// The current.
        /// </value>
        public override DataRow Record { get; set; }

        /// <summary>
        /// Gets the index of the current.
        /// </summary>
        /// <value>
        /// The index of the current.
        /// </value>
        public override int Index { get; set; }

        /// <summary>
        /// Gets or sets the field.
        /// </summary>
        /// <value>
        /// The field.
        /// </value>
        public override Field Field { get; set; }

        /// <summary>
        /// Gets or sets the data filter.
        /// </summary>
        /// <value>
        /// The data filter.
        /// </value>
        public override IDictionary<string, object> DataFilter { get; set; }

        /// <summary> Gets or sets the numeric. </summary>
        /// <value> The numeric. </value>
        public Numeric Numeric { get; set; }

        /// <summary> Sets the field. </summary>
        /// <param name = "field" > The field. </param>
        public void SetField( Field field )
        {
            try
            {
                Field = FormConfig.GetField( field );
            }
            catch( Exception ex )
            {
                Fail( ex );
            }
        }

        /// <summary> Sets the binding source. </summary>
        /// <typeparam name = "T1" > The type of the 1. </typeparam>
        /// <param name = "bindingsource" > The bindingsource. </param>
        public void SetDataSource<T1>( T1 bindingsource ) where T1 : IBindingList
        {
            try
            {
                if( bindingsource is BindingSource _binder
                    && _binder?.DataSource != null )
                {
                    try
                    {
                        DataSource = _binder.DataSource;
                    }
                    catch( Exception ex )
                    {
                        Fail( ex );
                    }
                }
            }
            catch( Exception ex )
            {
                Fail( ex );
            }
        }

        /// <summary> Sets the binding source. </summary>
        /// <typeparam name = "T1" > The type of the 1. </typeparam>
        /// <typeparam name = "T2" > The type of the 2. </typeparam>
        /// <param name = "bindinglist" > The bindingsource. </param>
        /// <param name = "dict" > The dictionary. </param>
        public void SetDataSource<T1, T2>( T1 bindinglist, T2 dict )
            where T1 : IBindingList where T2 : IDictionary<string, object>
        {
            try
            {
                if( dict?.Any() == true
                    && bindinglist is BindingSource _list )
                {
                    try
                    {
                        var _filter = string.Empty;

                        foreach( var _kvp in dict )
                        {
                            if( Verify.Input( _kvp.Key )
                                && _kvp.Value != null )
                            {
                                _filter += $"{_kvp.Key} = {_kvp.Value} AND";
                            }
                        }

                        if( _filter?.Length > 0
                            && _list?.DataSource != null )
                        {
                            DataSource = _list?.DataSource;
                            Filter = _filter?.TrimEnd( " AND".ToCharArray() );
                        }
                    }
                    catch( Exception ex )
                    {
                        Fail( ex );
                    }
                }
            }
            catch( Exception ex )
            {
                Fail( ex );
            }
        }

        /// <summary> Sets the binding source. </summary>
        /// <param name = "data" > The data. </param>
        public void SetDataSource( IEnumerable<object> data )
        {
            if( Verify.Input( data ) )
            {
                try
                {
                    DataSource = data?.ToList();
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }
        }

        /// <summary> Sets the binding source. </summary>
        /// <typeparam name = "T1" > The type of the 1. </typeparam>
        /// <param name = "data" > The data. </param>
        /// <param name = "dict" > The dictionary. </param>
        public void SetDataSource<T1>( IEnumerable<T1> data, IDictionary<string, object> dict )
            where T1 : IEnumerable<DataRow>
        {
            if( Verify.Sequence( data )
                && Verify.Map( dict ) )
            {
                try
                {
                    var _filter = string.Empty;

                    foreach( var _kvp in dict )
                    {
                        if( Verify.Input( _kvp.Key )
                            && _kvp.Value != null )
                        {
                            _filter += $"{_kvp.Key} = {_kvp.Value} AND";
                        }
                    }

                    DataSource = data?.ToList();
                    Filter = _filter?.TrimEnd( " AND".ToCharArray() );
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }
        }

        /// <summary> Sets the binding source. </summary>
        /// <typeparam name = "T1" > The type of the 1. </typeparam>
        /// <typeparam name = "T2" > The type of the 2. </typeparam>
        /// <typeparam name = "T3" > The type of the 3. </typeparam>
        /// <param name = "data" > The data. </param>
        /// <param name = "field" > The field. </param>
        /// <param name = "filter" > The dictionary. </param>
        public void SetDataSource<T1, T2, T3>( IEnumerable<T1> data, T2 field, T3 filter )
            where T1 : IEnumerable<DataRow> where T2 : struct
        {
            if( Verify.Sequence( data )
                && Validate.Field( field ) )
            {
                try
                {
                    if( Verify.Input( filter ) )
                    {
                        DataSource = data.ToList();
                        DataMember = field.ToString();
                        Filter = $"{field} = {filter}";
                    }
                    else
                    {
                        DataSource = data.ToList();
                        DataMember = field.ToString();
                    }
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }
        }

        /// <summary> Sets the binding source. </summary>
        /// <typeparam name = "T1" > The type of the 1. </typeparam>
        /// <param name = "data" > The data. </param>
        /// <param name = "field" > The field. </param>
        public void SetDataSource<T1>( IEnumerable<T1> data, object field = null )
            where T1 : IEnumerable<DataRow>
        {
            if( Verify.Sequence( data ) )
            {
                try
                {
                    if( Verify.Ref( field ) )
                    {
                        DataSource = data.ToList();
                        DataMember = field?.ToString();
                    }
                    else
                    {
                        DataSource = data.ToList();
                    }
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }
        }

        /// <summary> Sets the bindings. </summary>
        /// <typeparam name = "T1" > The type of the 1. </typeparam>
        /// <typeparam name = "T2" > The type of the 2. </typeparam>
        /// <param name = "data" > The data. </param>
        /// <param name = "dict" > The dictionary. </param>
        public void SetDataSource<T1, T2>( IEnumerable<T1> data, T2 dict )
            where T1 : IEnumerable<DataRow> where T2 : IDictionary<string, object>
        {
            if( Verify.Sequence( data )
                && Verify.Map( dict ) )
            {
                try
                {
                    var _filter = string.Empty;

                    foreach( var _kvp in dict )
                    {
                        if( Verify.Input( _kvp.Key )
                            && Verify.Ref( _kvp.Value ) )
                        {
                            _filter += $"{_kvp.Key} = {_kvp.Value} AND";
                        }
                    }

                    DataSource = data?.ToList();
                    Filter = _filter?.TrimEnd( " AND".ToCharArray() );
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }
        }

        /// <summary> Sets the binding source. </summary>
        /// <typeparam name = "T1" > The type of the 1. </typeparam>
        /// <typeparam name = "T2" > The type of the 2. </typeparam>
        /// <param name = "data" > The data. </param>
        /// <param name = "field" > The field. </param>
        /// <param name = "filter" > The filter. </param>
        public void SetDataSource<T1, T2>( IEnumerable<T1> data, T2 field, object filter = null )
            where T1 : IEnumerable<DataRow> where T2 : struct
        {
            if( Verify.Sequence( data )
                && Validate.Field( field ) )
            {
                try
                {
                    if( Verify.Ref( filter?.ToString() ) )
                    {
                        DataSource = data.ToList();
                        DataMember = field.ToString();
                        Filter = $"{field} = {filter}";
                    }
                    else
                    {
                        DataSource = data?.ToList();
                        DataMember = field.ToString();
                    }
                }
                catch( Exception ex )
                {
                    Fail( ex );
                }
            }
        }

        /// <summary>
        /// Gets the source.
        /// </summary>
        /// <returns>
        /// Returns the Source Enumeration
        /// </returns>
        public override Source GetSource()
        {
            try
            {
                return Validate.Source( Source )
                    ? Source
                    : default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default;
            }
        }

        /// <summary>
        /// Gets the field.
        /// </summary>
        /// <returns>
        /// </returns>
        public override Field GetField()
        {
            try
            {
                return Validate.Field( Field )
                    ? Field
                    : default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default;
            }
        }

        /// <summary>
        /// Gets the data filter.
        /// </summary>
        /// <returns>
        /// </returns>
        public override IDictionary<string, object> GetDataFilter()
        {
            try
            {
                return DataFilter?.Any() == true
                    ? DataFilter
                    : default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default;
            }
        }

        /// <summary>
        /// Gets the data table.
        /// </summary>
        /// <returns>
        /// </returns>
        public override DataTable GetDataTable()
        {
            try
            {
                return DataTable?.Rows?.Count > 0 && DataTable?.Columns?.Count > 0
                    ? DataTable
                    : default( DataTable );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default;
            }
        }

        /// <summary>
        /// Gets the data.
        /// </summary>
        /// <returns>
        /// </returns>
        public override IEnumerable<DataRow> GetData()
        {
            try
            {
                var _rows = DataTable?.AsEnumerable();

                return Verify.Input( _rows )
                    ? _rows
                    : default;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default;
            }
        }

        /// <summary>
        /// Gets the current row.
        /// </summary>
        /// <returns>
        /// </returns>
        public override DataRow GetCurrent()
        {
            try
            {
                return Record?.ItemArray?.Length > 0
                    ? Record
                    : default( DataRow );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default;
            }
        }

        /// <summary>
        /// Gets the index of the current.
        /// </summary>
        /// <returns>
        /// </returns>
        public override int GetIndex()
        {
            try
            {
                return Index > 0
                    ? Index
                    : -1;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default;
            }
        }

        /// <summary>
        /// Fails the specified ex.
        /// </summary>
        /// <param name="ex">The ex.</param>
        private protected override void Fail( Exception ex )
        {
            using var _error = new Error( ex );
            _error?.SetText();
            _error?.ShowDialog();
        }
    }
}