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

        private const string MYMETYPE_PDF = "application/pdf";
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
                // body.Contentは共用体。
                // Paragraph, SectionBreak, Table, TableOfContentsの
                // いずれかひとつが入っている
                if(element.Paragraph != null){
                    // paragraphの処理k
                    bodyText += GetParagraphText(element.Paragraph);
                } else if (element.SectionBreak != null){
                    // section breakの処理
                } else if (element.Table != null){
                    // Tableの処理
                } else if (element.TableOfContents != null){
                    // table of contentsの処理
                }
            }
            
            return bodyText;
       }

        private static string GetParagraphText(Paragraph paragraph)
        {
            string paragraphText = string.Empty;

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
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " start");
            DocumentsResource.GetRequest request = Service.Documents.Get(documentId);
            Document doc = request.Execute();
            Console.WriteLine("Found {0}", doc.Title);
            return doc;
        }

        // TODO: 指定フォルダ配下のファイル一括取得のメソッドを実装する
        public static void BatchReplace(string folderId, string origin, string changed)
        {
            var files = GoogleDrive.GoogleDriveInternal.ListFiles(folderId);
            var docs = files.Select(f => GoogleDocs.GoogleDocsInternal.GetDocument(f.Id));
            foreach(var doc in docs)
                ReplaceText(doc, origin, changed);
        }

        // public static string ExportToPdf(string documentId)
        // {
        //     return GoogleDrive.GoogleDriveInternal.ExportFile(documentId, MYMETYPE_PDF );
        // }
    }
}
