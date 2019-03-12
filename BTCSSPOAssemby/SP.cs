using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Btcs.IO;


namespace Btcs.SP
{
    /// <summary>
    /// Interact with SharePoint
    /// </summary>
    public class SP
    {
        #region CTor
        public SP(ClientContext ctx)
        {
            this.Context = ctx;
        }
        #endregion

        #region Props
        /// <summary>
        /// The clientcontext to execute query from
        /// </summary>
        public ClientContext Context { get; set; }

        /// <summary>
        /// The object to browse from a SPO Site
        /// </summary>
        public enum BaseTemplate
        {
            /// <summary>
            /// Lists object
            /// </summary>
            List,
            /// <summary>
            /// Library object
            /// </summary>
            Library
        }
        #endregion


        #region Public Methods
        /// <summary>
        /// Return a SP list from its title
        /// </summary>
        /// <param name="name">The list Title</param>
        /// <returns></returns>
        public List GetSPList(string name)
        {
            using (this.Context)
            {
                var result = this.Context.Web.GetListByTitle(name);
                this.Context.Load(result);
                this.Context.ExecuteQuery();
                return result;
            }

        }

        /// <summary>
        /// Return all lsits from a sharepoint site
        /// </summary>
        /// <returns></returns>
        public ListCollection GetAllSPList(bool showHidden = false)
        {
            using (this.Context)
            {
                ListCollection lists = this.Context.Web.Lists;

                if (showHidden)
                {
                    this.Context.Load(lists);
                }
                else
                {
                    this.Context.LoadQuery(lists.Where(l => l.Hidden == false));
                }
                Context.ExecuteQuery();
                return lists;
            }
        }

        /// <summary>
        /// Retrieve lists from a listcollection filtered by baseTemplate
        /// </summary>
        /// <param name="lists">The listCollection to process</param>
        /// <param name="objectType">The object type to filter</param>
        /// <param name="showHidden">Trigger to show hidden objects (default false)</param>
        /// <returns></returns>
        public List<List> GetListFromBaseTemplate(ListCollection lists, BaseTemplate objectType, bool showHidden = false)
        {
            var result = new List<List>();
            foreach (List list in lists)
            {
                if (showHidden)
                {
                    switch (objectType)
                    {
                        case BaseTemplate.List:
                            if (list.BaseTemplate == 100)
                                result.Add(list);
                            break;
                        case BaseTemplate.Library:
                            if (list.BaseTemplate == 101)
                                result.Add(list);
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    switch (objectType)
                    {
                        case BaseTemplate.List:
                            if (list.BaseTemplate == 100 && list.Hidden == false)
                                result.Add(list);
                            break;
                        case BaseTemplate.Library:
                            if (list.BaseTemplate == 101 && list.Hidden == false)
                                result.Add(list);
                            break;
                        default:
                            break;
                    }
                }

            }
            return result;
        }

        /// <summary>
        /// Retrieve all listitems in a list
        /// </summary>
        /// <param name="library">The list Name</param>
        /// <param name="rowlimit">The Row limit (default 100)</param>
        /// <returns>A list object of ListItems</returns>
        public List<ListItem> GetAllListItems(string listName, int rowLimit = 100)
        {
            using (this.Context)
            {
                List<ListItem> items = new List<ListItem>();

                List list = this.Context.Web.Lists.GetByTitle(listName);
                ListItemCollectionPosition position = null;
                var camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View Scope='RecursiveAll'>
                <Query>
                    <OrderBy Override='TRUE'><FieldRef Name='ID'/></OrderBy>
                </Query>
                <ViewFields>
                <FieldRef Name='Title'/><FieldRef Name='Modified' /><FieldRef Name='Editor' /><FieldRef Name='FileLeafRef' /><FieldRef Name='FileRef' /></ViewFields><RowLimit Paged='TRUE'>" + rowLimit + "</RowLimit></View>";

                do
                {
                    ListItemCollection listItems = null;
                    camlQuery.ListItemCollectionPosition = position;
                    listItems = list.GetItems(camlQuery);
                    this.Context.Load(listItems);
                    this.Context.ExecuteQuery();
                    position = listItems.ListItemCollectionPosition;
                    items.AddRange(listItems.ToList());
                }
                while (position != null);

                return items;
            }
        }

        /// <summary>
        /// Copy a file to a SharePoint library choosing between direct upload and slice by slice depending on the file size
        /// </summary>
        /// <param name="libraryName"></param>
        /// <param name="fileName"></param>
        /// <param name="itemNormalizedPath"></param>
        /// <param name="fileChunkSizeInMB"></param>
        public void UploadFileToSPO(string libraryName, string fileName, string itemNormalizedPath, int fileChunkSizeInMB = 3)
        {
            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();

            // Get the folder to upload into. 
            List docs = this.Context.Web.Lists.GetByTitle(libraryName);
            this.Context.Load(docs, l => l.RootFolder);
            // Get the information about the folder that will hold the file.
            this.Context.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            this.Context.ExecuteQuery();

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
                    this.Context.Load(uploadFile);
                    this.Context.ExecuteQuery();
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
                                        this.Context.ExecuteQuery();
                                        // We set fileoffset as the pointer where the next slice will be added
                                        fileoffset = bytesUploaded.Value;
                                    }
                                    first = false;
                                }
                            }
                            else
                            {
                                // We get a reference to our file
                                uploadFile = this.Context.Web.GetFileByServerRelativeUrl(itemNormalizedPath);

                                // We check if it is the last block
                                if (last)
                                {
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // We end the upload by calling FinishUpload
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        this.Context.ExecuteQuery();
                                    }
                                }
                                else // We continue the upload
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        this.Context.ExecuteQuery();
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
                }
            }
        }

        /// <summary>
        /// TODO : Pass a custom object with metadata
        /// </summary>
        /// <param name="list">The list object where to create the folder</param>
        /// <param name="folders">The foldertoupload object containing the folder metadata</param>
        /// <param name="batchsize">The number of item to process before executing the query to serverside</param>
        public void CreateFolderByBatch(List list, List<FolderToUpload> folders, int batchsize)
        {
            foreach (FolderToUpload folder in folders)
            {
                var myFolder = list.RootFolder.Folders.Add(folder.Url);
                ListItem listitemFolder = Context.Web.GetListItem(folder.Url);
                listitemFolder["Created"] = folder.Created;// UTC Date
                listitemFolder["Modified"] = folder.Modified;// UTC Date
                listitemFolder["Author"] = folder.Owner;
                listitemFolder.Update();
            }
        }


        #endregion


    }
}
