﻿// <copyright file = "_dataFile.cs" company = "Terry D. Eppler">
// Copyright (c) Terry D. Eppler. All rights reserved.
// </copyright>

namespace BudgetExecution
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;

    /// <summary>
    /// 
    /// </summary>
    /// <seealso cref="FileBase" />
    [ SuppressMessage( "ReSharper", "UseObjectOrCollectionInitializer" ) ]
    [ SuppressMessage( "ReSharper", "MemberCanBePrivate.Global" ) ]
    [ SuppressMessage( "ReSharper", "AssignNullToNotNullAttribute" ) ]
    public class DataFile : FileBase, IFile
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DataFile"/> class.
        /// </summary>
        public DataFile()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataFile"/> class.
        /// </summary>
        /// <param name="input">The input.</param>
        public DataFile( string input )
        {
            _path = new BudgetPath( input );
            _fileInfo = new FileInfo( _path.GetFullPath() );
            _fullName = _fileInfo.FullName;
            _hasParent = CheckParent();
            _length = _fileInfo.Length;
            _attributes = _fileInfo.Attributes;
            _security = _fileInfo.GetAccessControl();
            _creationDate = _fileInfo.CreationTime;
            _changedDate = _fileInfo.LastWriteTime;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataFile"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        public DataFile( IPath path )
        {
            _path = path;
            _fileInfo = new FileInfo( _path.GetFullPath() );
            _fullName = _fileInfo.FullName;
            _hasParent = CheckParent();
            _length = _fileInfo.Length;
            _attributes = _fileInfo.Attributes;
            _security = _fileInfo.GetAccessControl();
            _creationDate = _fileInfo.CreationTime;
            _changedDate = _fileInfo.LastWriteTime;
        }

        /// <summary>
        /// Creates the specified file path.
        /// </summary>
        /// <param name="filePath">
        /// The file path.
        /// </param>
        /// <returns></returns>
        public static FileInfo Create( string filePath )
        {
            try
            {
                return Verify.Input( filePath )
                    ? new FileInfo( filePath )
                    : default( FileInfo );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( FileInfo );
            }
        }

        /// <summary>
        /// Transfers the specified folder.
        /// </summary>
        /// <param name="folder">The folder.</param>
        public void Transfer( DirectoryInfo folder )
        {
            // Check if the target directory exists, if not, create it.
            if( !Directory.Exists( folder.FullName ) )
            {
                Directory.CreateDirectory( folder.FullName );
            }

            try
            {
                foreach( var fileinfo in folder?.GetFiles() )
                {
                    Directory.Move( fileinfo.FullName, folder.Name );
                }
            }
            catch( IOException ex )
            {
                Fail( ex );
            }
        }

        /// <summary>
        /// Determines whether this instance contains the object.
        /// </summary>
        /// <param name="search">The search.</param>
        /// <returns>
        ///   <c>true</c> if [contains] [the specified search]; otherwise, <c>false</c>.
        /// </returns>
        public bool Contains( string search )
        {
            try
            {
                var _input = _path?.GetFullPath();

                if( Verify.Input( _input )
                    && File.Exists( _input ) )
                {
                    using var _reader = new StreamReader( _input );
                    var _text = _reader?.ReadLine();
                    var _result = false;

                    while( _text == string.Empty )
                    {
                        if( Regex.IsMatch( _text, search ) )
                        {
                            _result = true;
                            break;
                        }

                        _text = _reader.ReadLine();
                    }

                    return _result;
                }

                return false;
            }
            catch( IOException ex )
            {
                Fail( ex );
                return false;
            }
        }

        /// <summary>
        /// Searches the specified pattern.
        /// </summary>
        /// <param name="pattern">The pattern.</param>
        /// <returns></returns>
        public IEnumerable<FileInfo> Search( string pattern )
        {
            if( Verify.Input( pattern ) )
            {
                try
                {
                    var _input = _path?.GetFullPath();

                    if( Verify.Input( _input )
                        && File.Exists( _input ) )
                    {
                        var _enumerable = Directory.EnumerateFiles( _input, pattern );
                        var _list = new List<FileInfo>();

                        foreach( var file in _enumerable )
                        {
                            _list.Add( new FileInfo( file ) );
                        }

                        return Verify.Input( _list )
                            ? _list
                            : default( List<FileInfo> );
                    }
                }
                catch( IOException ex )
                {
                    Fail( ex );
                    return default( IEnumerable<FileInfo> );
                }
            }

            return default( IEnumerable<FileInfo> );
        }

        /// <summary>Returns a string that
        /// represents the current object.
        /// </summary>
        /// <returns>A string that represents
        /// the current object.
        /// </returns>
        public override string ToString()
        {
            try
            {
                return Verify.Input( _fullName )
                    ? _fullName
                    : string.Empty;
            }
            catch( IOException ex )
            {
                Fail( ex );
                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <returns></returns>
        public string GetFullName()
        {
            try
            {
                return Verify.Input( _fullName )
                    ? _fullName
                    : default( string );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( string );
            }
        }

        /// <summary>
        /// Gets the data path.
        /// </summary>
        /// <returns></returns>
        public IPath GetDataPath()
        {
            try
            {
                return File.Exists( _path?.GetFullPath() )
                    ? _path
                    : default( IPath );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( IPath );
            }
        }

        /// <summary>
        /// Gets the extension.
        /// </summary>
        /// <returns></returns>
        public string GetFileExtension()
        {
            try
            {
                var _extension = _path?.GetFileExtension();

                return Verify.Input( _extension )
                    ? _extension
                    : string.Empty;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( string );
            }
        }

        /// <summary>
        /// Gets the full path.
        /// </summary>
        /// <returns></returns>
        public string GetFilePath()
        {
            try
            {
                var _input = _path?.GetFullPath();

                return Verify.Input( _input )
                    ? _input
                    : string.Empty;
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( string );
            }
        }

        /// <summary>
        /// Gets the parent.
        /// </summary>
        /// <returns></returns>
        public IFolder GetParentFolder()
        {
            try
            {
                return CheckParent()
                    ? new Folder( this )
                    : default( Folder );
            }
            catch( IOException ex )
            {
                Fail( ex );
                return default( IFolder );
            }
        }

        /// <summary>
        /// Browses this instance.
        /// </summary>
        /// <returns></returns>
        public static IFile Browse()
        {
            try
            {
                var _dialog = new OpenFileDialog();
                _dialog.CheckFileExists = true;
                _dialog.CheckPathExists = true;
                var _file = new DataFile( _dialog?.FileName );

                return File.Exists( _file?.GetFilePath() )
                    ? _file
                    : default( DataFile );
            }
            catch( Exception ex )
            {
                Fail( ex );
                return default( IFile );
            }
        }
    }
}
