using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace MsOneDriveRestApi
{
    public class MsOneDrive
    {
        const string serviceEndpoint = "https://graph.microsoft.com/v1.0/me/drive/";
        const string appRoot = "special/approot/";

        public static async Task<bool> IsAppRootExist()
        {
            if (await GetStringFromApiAsync(appRoot) != null)
            {
                return true;
            }

            return false;
        }

        static string BaseRoot(bool appRoot)
        {
            if (appRoot)
                return "special/approot";
            else
                return "root";
        }

        #region File

        #region Get

        /// <summary>
        /// Get file from path
        /// </summary>
        /// <param name="path">File name or path and file name (FloderA/TextB.txt)</param>
        /// <param name="fromAppRoot">App root is in root/Apps/Appname</param>
        /// <returns></returns>
        public static async Task<OneDriveFile> GetFile(string path, bool fromAppRoot = false)
        {
            string endpoint = $"{BaseRoot(fromAppRoot)}:/{path}";

            try
            {
                var json = await GetStringFromApiAsync(endpoint);
                return OneDriveItem.FromJson(json) as OneDriveFile;
            }
            catch (Exception)
            {
                throw;
            }
        }

        #endregion

        #region Upload

        /// <summary>
        /// Upload file to OneDrive AppRoot folder
        /// </summary>
        /// <param name="fileName">File name or path and file name (FloderA/TextB.txt)</param>
        /// <param name="file"></param>
        /// <param name="toAppRoot">App root is in root/Apps/Appname</param>
        /// <returns></returns>
        public static async Task<OneDriveFile> UploadFile(string fileName, Stream stream, bool toAppRoot = false)
        {
            string endpoint = $"{BaseRoot(toAppRoot)}:/{fileName}:/content";
            
            var streamContent = new StreamContent(stream, (int)stream.Length);
            streamContent.Headers.ContentType = System.Net.Http.Headers.MediaTypeHeaderValue.Parse("application/octet-stream");

            try
            {
                var json = await PutToApiAsync(endpoint, streamContent);
                return OneDriveItem.FromJson(json) as OneDriveFile;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                stream.Dispose();
            }
        }

        #endregion

        #region Delete

        public static async Task<bool> DeleteFileAsync(string id)
        {
            return await DeleteToApiAsync($"items/{id}");
        }

        #endregion

        #region Rename

        public static async Task<bool> RenameFileAsync(string fileId, string newName)
        {
            string endpoint = $"items/{fileId}";
            string reBody = $"{{'name': '{newName}' }}";
            var httpContent = new StringContent(reBody, Encoding.UTF8, "application/json");

            return await PatchToApiAsync(endpoint, httpContent);
        }

        #endregion

        #endregion

        #region Folder

        #region Get

        public static async Task<OneDriveFolder> GetFolder(string path, bool fromAppRoot = false)
        {
            string endpoint = $"{BaseRoot(fromAppRoot)}:/{path}";

            try
            {
                var json = await GetStringFromApiAsync(endpoint);
                return OneDriveItem.FromJson(json) as OneDriveFolder;
            }
            catch (Exception)
            {
                throw;
            }
        }

        #endregion

        #region Create

        /// <summary>
        /// Create folder on OneDrive
        /// </summary>
        /// <param name="name">Folder name</param>
        /// <param name="conflictBehavior"></param>
        /// <param name="inAppRoot">App root is in root/Apps/Appname</param>
        /// <returns></returns>
        public static async Task<OneDriveFolder> CreateFolder(string name, ConflictBehavior conflictBehavior, bool inAppRoot = false)
        {
            string folder = $"{{'name': '{name}', 'folder': {{}}, '@microsoft.graph.conflictBehavior': '{conflictBehavior.ToString().ToLower()}'}}";

            var content = new StringContent(folder, Encoding.UTF8, "application/json");

            try
            {
                string endpoint = $"{BaseRoot(inAppRoot)}/children";

                var json = await PostToApiAsync(endpoint, content);
                return OneDriveItem.FromJson(json) as OneDriveFolder;
            }
            catch (Exception)
            {
                throw;
            }
        }

        #endregion

        #region Delete

        public static async Task<bool> DeleteFolderAsync(string id)
        {
            return await DeleteToApiAsync($"items/{id}");
        }

        #endregion

        #region Rename

        public static async Task<bool> RenameFolderAsync(string folderId, string newName)
        {
            string endpoint = $"items/{folderId}";
            string reBody = $"{{'name': '{newName}' }}";
            var httpContent = new StringContent(reBody, Encoding.UTF8, "application/json");

            return await PatchToApiAsync(endpoint, httpContent);
        }

        #endregion

        #endregion

        #region Helper

        static async Task<bool> PatchToApiAsync(string endpoint, HttpContent httpContent)
        {
            var token = await MsGraph.GetTokenAsync();
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                Uri url = new Uri($"{serviceEndpoint}{endpoint}");

                var request = new HttpRequestMessage(new HttpMethod("PATCH"), url)
                {
                    Content = httpContent
                };


                var response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    return true;
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Conflict)
                {
                    throw new Exception(response.ReasonPhrase);
                }
            }

            return false;
        }

        static async Task<bool> DeleteToApiAsync(string endpoint)
        {
            bool del = false;

            var token = await MsGraph.GetTokenAsync();
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                Uri url = new Uri($"{serviceEndpoint}{endpoint}");

                var response = await client.DeleteAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    del = true;
                }
            }

            return del;
        }

        static async Task<string> PutToApiAsync(string endpoint, HttpContent httpContent)
        {
            var token = await MsGraph.GetTokenAsync();
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                Uri url = new Uri($"{serviceEndpoint}{endpoint}");

                var request = new HttpRequestMessage(HttpMethod.Put, url)
                {
                    Content = httpContent
                };


                var response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    return content;
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Conflict)
                {
                    throw new Exception(response.ReasonPhrase);
                }
            }

            return null;
        }

        static async Task<string> PostToApiAsync(string endpoint, HttpContent httpContent)
        {
            var token = await MsGraph.GetTokenAsync();
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                Uri url = new Uri($"{serviceEndpoint}{endpoint}");

                var response = await client.PostAsync(url, httpContent);

                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    return content;
                }
                else if (response.StatusCode == System.Net.HttpStatusCode.Conflict)
                {
                    throw new Exception(response.ReasonPhrase);
                }
            }

            return null;
        }

        static async Task<string> GetStringFromApiAsync(string endpoint)
        {
            try
            {
                var token = await MsGraph.GetTokenAsync();
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                    Uri url = new Uri($"{serviceEndpoint}{endpoint}");

                    var response = await client.GetAsync(url);

                    if (response.IsSuccessStatusCode)
                    {
                        return await response.Content.ReadAsStringAsync();
                    }
                }
            }
            catch (Exception)
            {

            }

            return null;
        }

        #endregion

        #region OneDriveItem

        public enum ConflictBehavior
        {
            Fail,
            Replace,
            Rename
        }

        public abstract class OneDriveItem
        {
            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("createdDateTime")]
            public DateTimeOffset CreatedDateTime { get; set; }

            [JsonProperty("id")]
            public string Id { get; set; }

            [JsonProperty("lastModifiedDateTime")]
            public DateTimeOffset LastModifiedDateTime { get; set; }

            [JsonProperty("parentReference")]
            public ParentReference ParentReference { get; set; }

            [JsonProperty("webUrl")]
            public string WebUrl { get; set; }

            [JsonIgnore]
            public bool IsFile { get; set; }

            public string ToJson()
            {
                if (IsFile)
                {
                    return JsonConvert.SerializeObject(this as OneDriveFile);
                }
                else
                {
                    return JsonConvert.SerializeObject(this as OneDriveFolder);
                }
            }

            public static OneDriveItem FromJson(string json)
            {

                var model = Newtonsoft.Json.Linq.JObject.Parse(json);

                if (model["folder"] != null)
                {
                    return JsonConvert.DeserializeObject<OneDriveFolder>(json);
                }
                else
                {
                    return JsonConvert.DeserializeObject<OneDriveFile>(json);
                }
            }
        }

        public class OneDriveFolder : OneDriveItem
        {
            [JsonProperty("folder")]
            public Folder Folder { get; set; }
        }

        public class OneDriveFile : OneDriveItem
        {
            public OneDriveFile()
            {
                IsFile = true;
            }

            [JsonProperty("size")]
            public int Size { get; set; }

            [JsonProperty("file")]
            public File File { get; set; }
        }

        public class File
        {
            [JsonProperty("mimeType")]
            public string MimeType { get; set; }
        }

        public class Folder
        {
            [JsonProperty("childCount")]
            public int ChildCount { get; set; }

            [JsonProperty("view")]
            public View View { get; set; }
        }

        public class View
        {
            [JsonProperty("viewType")]
            public string ViewType { get; set; }

            [JsonProperty("sortBy")]
            public string SortBy { get; set; }

            [JsonProperty("sortOrder")]
            public string SortOrder { get; set; }
        }

        public class ParentReference
        {
            [JsonProperty("driveId")]
            public string DriveId { get; set; }

            [JsonProperty("d")]
            public string Id { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("path")]
            public string Path { get; set; }
        }

        #endregion
    }
}
