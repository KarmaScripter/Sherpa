﻿// ******************************************************************************************
//     Assembly:                Budget Execution
//     Author:                  Terry D. Eppler
//     Created:                 04-29-2023
// 
//     Last Modified By:        Terry D. Eppler
//     Last Modified On:        05-31-2023
// ******************************************************************************************
// <copyright file="ThrowIf.cs" company="Terry D. Eppler">
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
//   ThrowIf.cs
// </summary>
// ******************************************************************************************

namespace BudgetExecution
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Data.Common;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using System.Net.Mail;
    using System.Windows.Forms;
    using OfficeOpenXml;

    /// <summary>
    /// 
    /// </summary>
    [ SuppressMessage( "ReSharper", "UnusedType.Global" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBeInternal" ) ]
    [ SuppressMessage( "ReSharper", "ConvertTypeCheckPatternToNullCheck" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "InternalOrPrivateMemberNotDocumented" ) ]
    public static class ThrowIf
    {
        /// <summary>
        /// Nulls the specified argument.
        /// </summary>
        /// <param name="argument">The argument.</param>
        /// <param name="paramName">Name of the parameter.</param>
        /// <exception cref="System.ArgumentNullException"></exception>
        public static void Null( object argument, string paramName )
        {
            switch( argument )
            {
                case string _args:
                {
                    if( string.IsNullOrEmpty( _args ) )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case Action _action:
                {
                    if( _action == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case Query _query:
                {
                    if( _query == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case DataBuilder _dataBuilder:
                {
                    if( _dataBuilder == null )
                    {
                        var _message = @$"The  '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case DataModel _dataModel:
                {
                    if( _dataModel == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case Form _form:
                {
                    if( _form == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case Control _control:
                {
                    if( _control == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case EventArgs _eventArgs:
                {
                    if( _eventArgs == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case DataConfig _dataConfig:
                {
                    if( _dataConfig == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case SqlConfig _sqlConfig:
                {
                    if( _sqlConfig == null )
                    {
                        var _message = @$"The SqlConfig '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case DataUnit _dataUnit:
                {
                    if( _dataUnit == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case Folder _folder:
                {
                    if( _folder == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case DataPath _path:
                {
                    if( _path == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case SqlStatement _sqlStatement:
                {
                    if( _sqlStatement == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case DataAdapter _adapter:
                {
                    if( _adapter == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case DataRow _row:
                {
                    var _items = _row?.ItemArray;
                    if( ( _items.Length == 0 ) 
                       || ( _items == null ) )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case EmailConfig _outlookConfig:
                {
                    if( _outlookConfig == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case EmailContent _emailContent:
                {
                    if( _emailContent == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case MailMessage _mailMessage:
                {
                    if( _mailMessage == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case ExcelWorksheet _worksheet:
                {
                    if( _worksheet == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case ExcelWorkbook _workbook:
                {
                    if( _workbook == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case ExcelPackage _package:
                {
                    if( _package == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case ExcelRange _range:
                {
                    if( _range == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case Grid _grid:
                {
                    if( _grid == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case ExcelCellBase _cell:
                {
                    if( _cell == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                default:
                {
                    if( argument == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
            }
        }

        /// <summary>
        /// Empty throws exception if 'argument' is null.
        /// </summary>
        /// <param name="argument">The argument.</param>
        /// <param name="paramName">Name of the parameter.</param>
        /// <exception cref="System.ArgumentException"></exception>
        /// <exception cref="System.ArgumentNullException">ArgumentNullException</exception>
        public static void Empty( object argument, string paramName )
        {
            switch( argument )
            {
                case string _args:
                {
                    if( string.IsNullOrEmpty( _args ) )
                    {
                        var _message = @$"The '{paramName}' is empty!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
                case IListSource _listSource:
                {
                    if( _listSource == null )
                    {
                        var _message = @$"The IListSource '{paramName}' is empty!";
                        throw new ArgumentException( _message );
                    }
                    
                    break;
                }
                case IBindingList _bindingList:
                {
                    if( _bindingList == null 
                       || _bindingList.IsEmpty( ) )
                    {
                        var _message = @$"The IBindingList '{paramName}' is empty!";
                        throw new ArgumentException( _message );
                    }

                    break;
                }
                case IEnumerable<string> _list:
                {
                    if( _list == null 
                       || _list.Any( ) != true )
                    {
                        var _message = @$"The IEnumerable<string> '{paramName}' is empty!";
                        throw new ArgumentException( _message );
                    }
                    
                    break;
                }
                case IEnumerable<byte> _data:
                {
                    if( _data == null
                       || _data.Any( ) != true )
                    {
                        var _message = @$"The IEnumerable<string> '{paramName}' is empty!";
                        throw new ArgumentException( _message );
                    }

                    break;
                }
                case IEnumerable<DataRow> _rows:
                {
                    if( _rows == null 
                       || _rows?.Any( ) != true )
                    {
                        var _message = @$"The IEnumerable<DataRow> '{paramName}' is empty!";
                        throw new ArgumentException( _message );
                    }
                    
                    break;
                }
                case ICollection _collection:
                {
                    if( _collection == null 
                       || _collection.IsEmpty( ) )
                    {
                        var _message = @$"The ICollection '{paramName}' is empty!";
                        throw new ArgumentException( _message );
                    }
                    
                    break;
                }
                case IDictionary<string, object> _dict:
                {
                    if( _dict == null 
                       || _dict.Keys.IsEmpty( ) )
                    {
                        var _message = @$"The IDictionary<string, object> '{paramName}' is empty!";
                        throw new ArgumentException( _message );
                    }
                    
                    break;
                }
                case IDictionary<string, string> _nvp:
                {
                    if( _nvp == null
                       || _nvp.Keys.IsEmpty( ) )
                    {
                        var _message = @$"The IDictionary<string, string> '{paramName}' is empty!";
                        throw new ArgumentException( _message );
                    }

                    break;
                }
                default:
                {
                    if( argument == null )
                    {
                        var _message = @$"The '{paramName}' is null!";
                        throw new ArgumentNullException( _message );
                    }

                    break;
                }
            }
        }

        /// <summary>
        /// Determines whether the specified argument is negative.
        /// </summary>
        /// <param name="argument">The argument.</param>
        /// <param name="paramName">The argument's name.</param>
        /// <exception cref="System.ArgumentOutOfRangeException"></exception>
        public static void NegativeOrZero( object argument, string paramName )
        {
            var _message = @$"{paramName} must be greater than zero";
            switch( argument )
            {
                case sbyte _sbyte:
                {
                    if( _sbyte <= 0 )
                    {
                        throw new ArgumentOutOfRangeException( paramName, _message );
                    }

                    break;
                }
                case byte _byte:
                {
                    if( _byte <= 0 )
                    {
                        throw new ArgumentOutOfRangeException( paramName, _message );
                    }

                    break;
                }
                case int _count:
                {
                    if( _count <= 0 )
                    {
                        throw new ArgumentOutOfRangeException( paramName, _message );
                    }

                    break;
                }
                case short _shorty:
                {
                    if( _shorty <= 0 )
                    {
                        throw new ArgumentOutOfRangeException( paramName, _message );
                    }

                    break;
                }
                case double _number:
                {
                    if( _number <= 0 )
                    {
                        throw new ArgumentOutOfRangeException( paramName, _message );
                    }

                    break;
                }
                case float _floaty:
                {
                    if( _floaty <= 0 )
                    {
                        throw new ArgumentOutOfRangeException( paramName, _message );
                    }

                    break;
                }
                case decimal _cash:
                {
                    if( _cash <= 0 )
                    {
                        throw new ArgumentOutOfRangeException( paramName, _message );
                    }

                    break;
                }
            }
        }
    }
}