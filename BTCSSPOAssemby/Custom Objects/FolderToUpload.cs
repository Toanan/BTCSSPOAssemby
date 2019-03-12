using System;

namespace Btcs.IO
{

    /// <summary>
    /// Object representing the folder to upload
    /// </summary>
    public class FolderToUpload
    {
        /// <summary>
        /// The folder target Url <site>/<library>/<folderNormalizedUrl>
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// The folder creation date
        /// </summary>
        public DateTime Created { get; set; }

        /// <summary>
        /// The folder modification date
        /// </summary>
        public DateTime Modified { get; set; }

        /// <summary>
        /// The folder Owner
        /// </summary>
        public string Owner { get; set; }
    }
}
