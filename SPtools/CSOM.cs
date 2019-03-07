using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;

namespace SPtools
{
    class CSOM
    {

        /// <summary>
        /// Retrieve all list from a SharePoint Online site
        /// </summary>
        /// <param name="Context">The SharePoint Online authenticated ClientContext</param>
        /// <returns>A Web object</returns>
        public static ListCollection getAllList(ClientContext Context)
        {
            ListCollection Libraries = Context.Web.Lists;
            Context.Load(Libraries);
            Context.ExecuteQuery();
            return Libraries;
        }

        /// <summary>
        /// Retrieve the Web object
        /// </summary>
        /// <param name="Context">The SharePoint Online authenticated ClientContext</param>
        /// <returns>A Web object</returns>
        public static Web GetWeb(ClientContext Context)
        {
            Web site = Context.Web;
            Context.Load(site);//, s => s.Url);
            Context.ExecuteQuery();
            return site;
        }

        /// <summary>
        /// Delete a field
        /// </summary>
        /// <param name="list">The list where the column is located</param>
        /// <param name="column">The column name to delete</param>
        /// <param name="Context">The ClientContext</param>
        /// <returns></returns>
        public static bool deleteField(List list, string column, ClientContext Context)
        {

            Field field = list.Fields.GetByInternalNameOrTitle(column);
            Context.Load(field);
            field.DeleteObject();
            list.Update();
            Context.ExecuteQuery();
            return true;
        }

        /// <summary>
        /// Retrieve all listitems in a list
        /// </summary>
        /// <param name="libName">The list</param>
        /// <param name="Context">The ClientContext</param>
        /// <returns></returns>
        public static List<ListItem> GetAllListItem(string listName, ClientContext Context, int Rows)
        {
            /*
            List<ListItem> items = new List<ListItem>();
            Context.Load(Context.Web, a => a.Lists);
            Context.ExecuteQuery();
            */
            List list = Context.Web.Lists.GetByTitle(listName);
            ListItemCollectionPosition position = null;
            var camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"<View Scope='RecursiveAll'>
            <Query>
                <OrderBy Override='TRUE'><FieldRef Name='ID'/></OrderBy>
            </Query>
            <ViewFields>
            <FieldRef Name='Title'/><FieldRef Name='Modified' /><FieldRef Name='Editor' /><FieldRef Name='FileLeafRef' /><FieldRef Name='FileRef' /></ViewFields><RowLimit Paged='TRUE'>" + Rows + "</RowLimit></View>";
            do
            {
                ListItemCollection listItems = null;
                camlQuery.ListItemCollectionPosition = position;
                listItems = list.GetItems(camlQuery);
                Context.Load(listItems);
                Context.ExecuteQuery();
                position = listItems.ListItemCollectionPosition;
                items.AddRange(listItems.ToList());
            }
            while (position != null);

            return items;
        }

        /// <summary>
        /// Copy file using file.add or StartUpload depending on file size
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="libraryName"></param>
        /// <param name="fileName"></param>
        /// <param name="fileChunkSizeInMB"></param>
        public void UploadFile(ClientContext ctx, string libraryName, string fileName, string itemNormalizedPath, int fileChunkSizeInMB = 3)
        {
            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();

            // Get the folder to upload into. 
            List docs = ctx.Web.Lists.GetByTitle(libraryName);
            ctx.Load(docs, l => l.RootFolder);
            // Get the information about the folder that will hold the file.
            ctx.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();

            // We create the file object
            Microsoft.SharePoint.Client.File uploadFile;

            // We calculate block size in bytes
            int blockSize = fileChunkSizeInMB * 1024 * 1024;

            // We retrieve the size of the file
            long fileSize = new FileInfo(fileName).Length;

            //If local file size < block size
            if (fileSize <= blockSize)
            {
                // We use File.add method to upload
                using (FileStream fs = new FileStream(fileName, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = itemNormalizedPath;
                    fileInfo.Overwrite = true;
                    uploadFile = docs.RootFolder.Files.Add(fileInfo);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();
                }
            }
            else
            {
                // We use the large file method
                ClientResult<long> bytesUploaded = null;

                FileStream fs = null;
                try
                {
                    fs = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using (BinaryReader br = new BinaryReader(fs))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // We read the local file by block 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // We check if we read the last block 
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            // We check if we read the first block 
                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // We add an empty file
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = itemNormalizedPath;
                                    fileInfo.Overwrite = true;
                                    uploadFile = docs.RootFolder.Files.Add(fileInfo);

                                    // We start upload by uploading the first block 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first block
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        ctx.ExecuteQuery();
                                        // We set fileoffset as the pointer where the next slice will be added
                                        fileoffset = bytesUploaded.Value;
                                    }
                                    first = false;
                                }
                            }
                            else
                            {
                                // We get a reference to our file
                                uploadFile = ctx.Web.GetFileByServerRelativeUrl(itemNormalizedPath);

                                // We check if it is the last block
                                if (last)
                                {
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // We end the upload by calling FinishUpload
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                    }
                                }
                                else // We continue the upload
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // Update fileoffset for the next block.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        }
                    }
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Dispose();
                    }
                    return uploadFile;
                }
            }
        }

    }
}
