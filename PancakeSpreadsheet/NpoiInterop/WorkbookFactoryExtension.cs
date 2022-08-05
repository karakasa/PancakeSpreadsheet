/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at
       http://www.apache.org/licenses/LICENSE-2.0
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

/*
 * This file is a derived work, copied and adapted from the project nissl-lab/npoi (https://github.com/nissl-lab/npoi)
 * to extend WorkbookFactory.Create support to seekable XSSF/HSSF streams with/without a password.
*/

using NPOI.HSSF.Record.Crypto;
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace NPOI.SS.UserModel
{
    public static class WorkbookFactoryExtension
    {
        public static IWorkbook Create(Stream stream, string password = "")
        {
            if (!stream.CanRead || !stream.CanSeek)
                throw new ArgumentException($"{nameof(stream)} is not seekable nor readable.");

            stream.Position = 0L;

            if (POIFSFileSystem.HasPOIFSHeader(stream))
            {
                stream.Position = 0L;

                if (string.IsNullOrEmpty(password))
                {
                    return new HSSFWorkbook(stream);
                }
                else
                {
                    var nPOIFSFileSystem = new NPOIFSFileSystem(stream);
                    try
                    {
                        return CreateFromOLE2FS(nPOIFSFileSystem, password);
                    }
                    finally
                    {
                        nPOIFSFileSystem?.Close();
                    }
                }
            }

            stream.Position = 0L;
            if (DocumentFactoryHelper.HasOOXMLHeader(stream))
            {
                if (!string.IsNullOrEmpty(password))
                {
                    throw new NotSupportedException("Password is not allowed for OOXML files. You should clear the password input.");
                }

                stream.Position = 0L;

                return new XSSFWorkbook(stream);
            }

            throw new NotSupportedException("Unknown file format.");
        }
        private static IWorkbook CreateFromOLE2FS(NPOIFSFileSystem fs, string password)
        {
            var root = fs.Root;
            if (root.HasEntry(Decryptor.DEFAULT_POIFS_ENTRY))
            {
                return new XSSFWorkbook(OPCPackage.Open(DocumentFactoryHelper.GetDecryptedStream(fs, password)));
            }

            if (password != null)
            {
                Biff8EncryptionKey.CurrentUserPassword = password;
            }

            try
            {
                return new HSSFWorkbook(root, preserveNodes: true);
            }
            finally
            {
                Biff8EncryptionKey.CurrentUserPassword = null;
            }
        }
        public static void SaveWithPassword(this IWorkbook workbook, Stream targetStream, string password = "")
        {
            if (string.IsNullOrEmpty(password))
            {
                workbook.Write(targetStream);
                return;
            }
        }
    }
}
