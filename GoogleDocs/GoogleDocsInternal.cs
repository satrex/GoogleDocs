using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Reflection;

namespace Satrex.GoogleDocs
{
    /// Manipulates GSuite internal document only by system.
    /// This class doesn't provide user document.
    public class GoogleDocsInternal
    {
        static string[] Scopes = { DocsService.Scope.Documents, 
        DocsService.Scope.Drive, DocsService.Scope.DriveFile};
        static string ApplicationName = "Google Docs Manipulator";
        private static DocsService service;

        public const string MYMETYPE_PDF = "application/pdf";
        
        public static string GOOGLE_MYMETYPE_FOLDER = @"application/vnd.google-apps.folder";
        private static string credentialFilePath = @"Secrets/client_secret.json";
        public static string CredentialFilePath
        {
            get { return credentialFilePath; }
            set { credentialFilePath = value; }
        }
        
        public static DocsService Service
        {
            get
            {
                if (service == null)
                {
                    service = CreateDocsService(CredentialFilePath);
                }
                return service;
            }
        }
        public static string GetExecutingDirectoryName()
        {
            var location = new Uri(Assembly.GetExecutingAssembly().GetName().CodeBase);
            var locationPath = location.AbsolutePath + location.Fragment;
            var locationDir = new FileInfo(locationPath).Directory;
            return locationDir.FullName;
        }        
        
        private static DocsService CreateDocsService(string credentialFilePath)
        {
            UserCredential credential;
            //認証プロセス。credPathが作成されていないとBrowserが起動して認証ページが開くので認証を行って先に進む
            var exeDir = GetExecutingDirectoryName();
            var secretFilePath = Path.Combine(exeDir, credentialFilePath);
            using (var stream = new FileStream(secretFilePath, FileMode.Open, FileAccess.Read))
            {
                string credPath = Path.Combine
                    (System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal),
                     ".credentials/docs.googleapis.com-satrex.json");
                //CredentialファイルがcredPathに保存される
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync
                    (GoogleClientSecrets.Load(stream).Secrets, 
                    Scopes, 
                    "user", 
                    CancellationToken.None,
                     new FileDataStore(credPath, true)).Result;
            }
            //API serviceを作成、Requestパラメータを設定
            var service = new DocsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }
        public GoogleDocsInternal()
        {
            
        }

        public static string GetTitle(string documentId)
        {
            DocumentsResource.GetRequest request = Service.Documents.Get(documentId);

            // Prints the title of the requested doc:
            // https://docs.google.com/document/d/195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE/edit
            Document doc = request.Execute();
            Console.WriteLine("The title of the doc is: {0}", doc.Title);
            return doc.Title;
        }

        public static string GetAllText(string documentId)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " start");
            DocumentsResource.GetRequest request = Service.Documents.Get(documentId);

            // Prints the title of the requested doc:
            // https://docs.google.com/document/d/195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE/edit
            Document doc = request.Execute();
            Console.WriteLine("The title of the doc is: {0}", doc.Title);

            var bodyText = string.Empty;
            foreach(var element in doc.Body.Content){
                bodyText += GetContent(element);
            }
            
            return bodyText;
       }

        public static string GetContent(StructuralElement element)
        {
            // body.Contentは共用体。
            // Paragraph, SectionBreak, Table, TableOfContentsの
            // いずれかひとつが入っている
            if (element.Paragraph != null)
            {
                // paragraphの処理k
                return GetParagraphText(element.Paragraph);
            }
            else if (element.SectionBreak != null)
            {
                // section breakの処理
            }
            else if (element.Table != null)
            {
                // Tableの処理
            }
            else if (element.TableOfContents != null)
            {
                // table of contentsの処理
            }
            return string.Empty;
        }

        public static string GetParagraphText(Paragraph paragraph)
        {
            string paragraphText = string.Empty;
            if (paragraph == null) return paragraphText;
            
            foreach(var element in paragraph.Elements)
            {
                paragraphText += element.TextRun.Content;
            }

            return paragraphText;
        }

        public static Document ReplaceText(Document document, string origin, string changed)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " start");
            return ReplaceText(document, origin, changed, true);
        }

        public async static Task<Document> ReplaceTextAsync(Document document, string origin, string changed)
        {
            return await ReplaceTextAsync(document, origin, changed, true);
        }

        public async static Task<Document> ReplaceTextAsync(Document document, string origin, string changed, bool isCaseSensitive)
        {
            var req = new BatchUpdateDocumentRequest();
            var replace =  new Request();
            replace.ReplaceAllText = new ReplaceAllTextRequest()
            {
                ContainsText = new SubstringMatchCriteria(){
                    Text = origin,
                    MatchCase = isCaseSensitive
                },
                ReplaceText = changed
            };

            req.Requests = new List<Request>();
            req.Requests.Add(replace);
            var updater = Service.Documents.BatchUpdate(req, document.DocumentId);
            var response = await updater.ExecuteAsync();
            return document;
        }

        
        /// <summary>
        /// 対象ファイル内の文字列を置換します。
        /// </summary>
        /// <param name="document">対象ファイルを指定します。</param>
        /// <param name="origin">置換される文字列を指定します。</param>
        /// <param name="changed">置換後の文字列を指定します。</param>
        /// <param name="isCaseSensitive">大文字小文字を区別する場合はtrue、それ以外はfalseを指定します。</param>
        /// <returns></returns>
        public static Document ReplaceText(Document document, string origin, string changed, bool isCaseSensitive)
        {
            var req = new BatchUpdateDocumentRequest();
            var replace =  new Request();
            replace.ReplaceAllText = new ReplaceAllTextRequest()
            {
                ContainsText = new SubstringMatchCriteria(){
                    Text = origin,
                    MatchCase = isCaseSensitive
                },
                ReplaceText = changed
            };

            req.Requests = new List<Request>();
            req.Requests.Add(replace);
            var updater = Service.Documents.BatchUpdate(req, document.DocumentId);
            var response = updater.Execute();
            return document;
        }

        public static Document GetDocument(string documentId)
        {
            // Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " start");
            // Console.WriteLine($"documentId:{documentId}");
            try
            {
                DocumentsResource.GetRequest request = Service.Documents.Get(documentId);
                // Console.WriteLine(request.PrettyPrint);
                Document doc = request.Execute();
                // Console.WriteLine("Found {0}", doc.Title);
                return doc;
            }
            catch (Google.GoogleApiException ex)
            {
                Console.Error.WriteLine("Googleドキュメントが取得できませんでした。MymeTypeを確認してください。");
                Console.Error.WriteLine($"{ex.GetType().Name} {ex.Message} ");
                throw ex;
            }
        }

        /// <summary>
        /// 指定フォルダ配下のファイルすべてに、文字列置換を行います。
        /// </summary>
        /// <param name="folderId">最上位のフォルダIDを指定します。</param>
        /// <param name="origin">置換される文字列を指定します。</param>
        /// <param name="changed">置換後の文字列を指定します。</param>
        public static void BatchReplace(string folderId, string origin, string changed)
        {
            var files = GoogleDrive.GoogleDriveInternal.ListFiles(folderId);
            var docs = files.Select(f => GoogleDocs.GoogleDocsInternal.GetDocument(f.Id));
            foreach(var doc in docs)
                ReplaceText(doc, origin, changed);
        }

        /// <summary>
        /// 指定フォルダ配下のファイルすべてに、文字列置換を行います。
        /// サブフォルダすべてを再帰的に処理します。
        /// </summary>
        /// <param name="folderId">最上位のフォルダIDを指定します。</param>
        /// <param name="origin">置換される文字列を指定します。</param>
        /// <param name="changed">置換後の文字列を指定します。</param>
        public static void BatchReplaceRecursively(string folderId, string origin, string changed, Action<string, string> onReplaced)
        {
            var files = GoogleDrive.GoogleDriveInternal.ListFiles(folderId);
            var docs = files.Where(f => f.MimeType == Satrex.GoogleDrive.GoogleDriveInternal.GOOGLE_MYMETYPE_DOCS)
                        .Select(f =>
                        {
                            System.Threading.Thread.Sleep(1100);
                            return GoogleDocs.GoogleDocsInternal.GetDocument(f.Id);
                        });
            foreach(var doc in docs)
            {
                System.Threading.Thread.Sleep(1100);
                ReplaceText(doc, origin, changed);
                if(onReplaced != null)
                    onReplaced(doc.DocumentId, doc.Title);
            }

            var folders = files.Where(f => f.MimeType == GoogleDrive.GoogleDriveInternal.GOOGLE_MYMETYPE_FOLDER);
            foreach(var folder in folders)
            {
                BatchReplaceRecursively(folder.Id, origin, changed, onReplaced);
            }
        }


        // public static string ExportToPdf(string documentId)
        // {
        //     return GoogleDrive.GoogleDriveInternal.ExportFile(documentId, MYMETYPE_PDF );
        // }
    }
}
