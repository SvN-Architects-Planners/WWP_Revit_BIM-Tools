
using Autodesk.Revit.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using Microsoft.Win32;
using System.Runtime.InteropServices;

public class Script
{
    public static void Execute(UIApplication uiapp)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        Window window = UiLoader.LoadWindow("ui.xaml");
        if (window == null)
        {
            TaskDialog.Show("WWP BIM Tools", "ui.xaml not found. Ensure ui.xaml is next to script.cs.");
            return;
        }

        new AccDocsSyncController(window);
        window.ShowDialog();
    }
}

internal static class UiLoader
{
    public static Window LoadWindow(string fileName)
    {
        string xamlPath = FindXamlPath(fileName);
        if (string.IsNullOrEmpty(xamlPath) || !File.Exists(xamlPath))
            return null;

        using (FileStream stream = new FileStream(xamlPath, FileMode.Open, FileAccess.Read))
        {
            return (Window)XamlReader.Load(stream);
        }
    }

    private static string FindXamlPath(string fileName)
    {
        List<string> candidates = new List<string>();

        try
        {
            candidates.Add(Path.Combine(Environment.CurrentDirectory, fileName));
        }
        catch { }

        try
        {
            string asmDir = Path.GetDirectoryName(typeof(UiLoader).Assembly.Location);
            if (!string.IsNullOrEmpty(asmDir))
                candidates.Add(Path.Combine(asmDir, fileName));
        }
        catch { }

        try
        {
            string procDir = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            if (!string.IsNullOrEmpty(procDir))
                candidates.Add(Path.Combine(procDir, fileName));
        }
        catch { }

        foreach (string candidate in candidates.Distinct())
        {
            if (File.Exists(candidate))
                return candidate;
        }

        try
        {
            string cwd = Environment.CurrentDirectory;
            for (int i = 0; i < 6; i++)
            {
                string guess = Path.Combine(
                    cwd,
                    "WWP_Revit_BIM-Tools.extension",
                    "WWP BIM Tools.tab",
                    "Cloud.panel",
                    "ACC Docs Attribute Sync.pushbutton",
                    fileName);
                if (File.Exists(guess))
                    return guess;

                DirectoryInfo parent = Directory.GetParent(cwd);
                if (parent == null)
                    break;
                cwd = parent.FullName;
            }
        }
        catch { }

        return null;
    }
}

internal class AccDocsSyncController
{
    private const string RedirectUri = "http://127.0.0.1:8765/callback/";
    private const string OAuthAuthorizeUrl = "https://developer.api.autodesk.com/authentication/v2/authorize";
    private const string OAuthTokenUrl = "https://developer.api.autodesk.com/authentication/v2/token";
    private const string DefaultScopes = "data:read data:write account:read";

    private readonly Window _window;
    private readonly TextBox _clientIdBox;
    private readonly PasswordBox _clientSecretBox;
    private readonly TextBox _redirectUriBox;
    private readonly Button _loginButton;
    private readonly ComboBox _hubCombo;
    private readonly ComboBox _projectCombo;
    private readonly TreeView _folderTree;
    private readonly Button _refreshFoldersButton;
    private readonly ListView _fileList;
    private readonly TextBox _excelPathBox;
    private readonly TextBlock _excelStatusText;
    private readonly Button _browseExcelButton;
    private readonly Button _applyButton;
    private readonly TextBox _logBox;
    private readonly TextBlock _statusText;

    private readonly ObservableCollection<FolderNode> _folderNodes = new ObservableCollection<FolderNode>();
    private readonly ObservableCollection<FileItem> _fileItems = new ObservableCollection<FileItem>();

    private AccAuthSession _session;
    private AccAuthClient _authClient;
    private AccDataClient _dataClient;
    private Dictionary<string, ExcelRow> _excelRows;

    public AccDocsSyncController(Window window)
    {
        _window = window;
        _clientIdBox = (TextBox)window.FindName("ClientIdBox");
        _clientSecretBox = (PasswordBox)window.FindName("ClientSecretBox");
        _redirectUriBox = (TextBox)window.FindName("RedirectUriBox");
        _loginButton = (Button)window.FindName("LoginButton");
        _hubCombo = (ComboBox)window.FindName("HubCombo");
        _projectCombo = (ComboBox)window.FindName("ProjectCombo");
        _folderTree = (TreeView)window.FindName("FolderTree");
        _refreshFoldersButton = (Button)window.FindName("RefreshFoldersButton");
        _fileList = (ListView)window.FindName("FileList");
        _excelPathBox = (TextBox)window.FindName("ExcelPathBox");
        _excelStatusText = (TextBlock)window.FindName("ExcelStatusText");
        _browseExcelButton = (Button)window.FindName("BrowseExcelButton");
        _applyButton = (Button)window.FindName("ApplyButton");
        _logBox = (TextBox)window.FindName("LogBox");
        _statusText = (TextBlock)window.FindName("StatusText");

        _redirectUriBox.Text = RedirectUri;

        _hubCombo.IsEnabled = false;
        _projectCombo.IsEnabled = false;
        _folderTree.IsEnabled = false;
        _refreshFoldersButton.IsEnabled = false;
        _browseExcelButton.IsEnabled = false;
        _applyButton.IsEnabled = false;

        _folderTree.ItemsSource = _folderNodes;
        _fileList.ItemsSource = _fileItems;

        _loginButton.Click += async (s, e) => await SignInAsync();
        _hubCombo.SelectionChanged += async (s, e) => await OnHubChangedAsync();
        _projectCombo.SelectionChanged += async (s, e) => await OnProjectChangedAsync();
        _folderTree.SelectedItemChanged += async (s, e) => await OnFolderSelectedAsync(e.NewValue as FolderNode);
        _folderTree.AddHandler(TreeViewItem.ExpandedEvent, new RoutedEventHandler(OnFolderExpanded));
        _refreshFoldersButton.Click += async (s, e) => await LoadTopFoldersAsync();
        _browseExcelButton.Click += (s, e) => BrowseExcel();
        _applyButton.Click += async (s, e) => await ApplyDescriptionsAsync();

        Log("Ready.");
    }

    private async Task SignInAsync()
    {
        string clientId = _clientIdBox.Text.Trim();
        string clientSecret = _clientSecretBox.Password.Trim();

        if (string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(clientSecret))
        {
            MessageBox.Show("Client ID and Client Secret are required.", "ACC Docs", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        _loginButton.IsEnabled = false;
        _statusText.Text = "Signing in...";

        try
        {
            _authClient = new AccAuthClient(OAuthAuthorizeUrl, OAuthTokenUrl, RedirectUri, DefaultScopes, Log);
            _session = await _authClient.AuthenticateAsync(clientId, clientSecret);
            _session.ClientId = clientId;
            _session.ClientSecret = clientSecret;
            _dataClient = new AccDataClient(_authClient, _session, Log);

            _statusText.Text = "Signed in";
            _hubCombo.IsEnabled = true;
            _projectCombo.IsEnabled = true;
            _folderTree.IsEnabled = true;
            _refreshFoldersButton.IsEnabled = true;
            _browseExcelButton.IsEnabled = true;
            _applyButton.IsEnabled = true;

            await LoadHubsAsync();
        }
        catch (Exception ex)
        {
            _statusText.Text = "Sign in failed";
            Log("Sign in failed: " + ex.Message);
        }
        finally
        {
            _loginButton.IsEnabled = true;
        }
    }

    private async Task LoadHubsAsync()
    {
        try
        {
            _statusText.Text = "Loading hubs...";
            List<HubInfo> hubs = await _dataClient.GetHubsAsync();
            _hubCombo.ItemsSource = hubs;
            _hubCombo.DisplayMemberPath = "Name";
            _hubCombo.SelectedValuePath = "Id";
            if (hubs.Count > 0)
                _hubCombo.SelectedIndex = 0;
            _statusText.Text = "Hubs loaded";
        }
        catch (Exception ex)
        {
            Log("Failed to load hubs: " + ex.Message);
            _statusText.Text = "Hub load failed";
        }
    }

    private async Task OnHubChangedAsync()
    {
        HubInfo hub = _hubCombo.SelectedItem as HubInfo;
        if (hub == null)
            return;

        try
        {
            _statusText.Text = "Loading projects...";
            List<ProjectInfo> projects = await _dataClient.GetProjectsAsync(hub.Id);
            _projectCombo.ItemsSource = projects;
            _projectCombo.DisplayMemberPath = "Name";
            _projectCombo.SelectedValuePath = "Id";
            if (projects.Count > 0)
                _projectCombo.SelectedIndex = 0;
            _statusText.Text = "Projects loaded";
        }
        catch (Exception ex)
        {
            Log("Failed to load projects: " + ex.Message);
            _statusText.Text = "Project load failed";
        }
    }
    private async Task OnProjectChangedAsync()
    {
        await LoadTopFoldersAsync();
    }

    private async Task LoadTopFoldersAsync()
    {
        HubInfo hub = _hubCombo.SelectedItem as HubInfo;
        ProjectInfo project = _projectCombo.SelectedItem as ProjectInfo;
        if (hub == null || project == null)
            return;

        try
        {
            _statusText.Text = "Loading folders...";
            _folderNodes.Clear();
            _fileItems.Clear();

            List<FolderNode> topFolders = await _dataClient.GetTopFoldersAsync(hub.Id, project.Id);
            foreach (FolderNode folder in topFolders)
                _folderNodes.Add(folder);

            _statusText.Text = "Folders loaded";
        }
        catch (Exception ex)
        {
            Log("Failed to load folders: " + ex.Message);
            _statusText.Text = "Folder load failed";
        }
    }
    private async void OnFolderExpanded(object sender, RoutedEventArgs e)
    {
        TreeViewItem item = e.OriginalSource as TreeViewItem;
        if (item == null)
            return;

        FolderNode node = item.DataContext as FolderNode;
        if (node == null || node.IsLoaded || node.IsPlaceholder)
            return;

        ProjectInfo project = _projectCombo.SelectedItem as ProjectInfo;
        if (project == null)
            return;

        try
        {
            node.IsLoading = true;
            List<FolderNode> children = await _dataClient.GetFolderChildrenAsync(project.Id, node.Id);
            node.Children.Clear();
            foreach (FolderNode child in children)
                node.Children.Add(child);
            node.IsLoaded = true;
        }
        catch (Exception ex)
        {
            Log("Failed to expand folder: " + ex.Message);
        }
        finally
        {
            node.IsLoading = false;
        }
    }

    private async Task OnFolderSelectedAsync(FolderNode node)
    {
        if (node == null)
            return;

        ProjectInfo project = _projectCombo.SelectedItem as ProjectInfo;
        if (project == null)
            return;

        try
        {
            _statusText.Text = "Loading files...";
            _fileItems.Clear();
            List<FileItem> files = await _dataClient.GetFilesInFolderAsync(project.Id, node.Id);
            foreach (FileItem file in files)
                _fileItems.Add(file);
            _statusText.Text = "Files loaded";
        }
        catch (Exception ex)
        {
            Log("Failed to load files: " + ex.Message);
            _statusText.Text = "File load failed";
        }
    }
    private void BrowseExcel()
    {
        OpenFileDialog dialog = new OpenFileDialog();
        dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls";
        dialog.Multiselect = false;

        if (dialog.ShowDialog() != true)
            return;

        _excelPathBox.Text = dialog.FileName;
        try
        {
            _excelRows = ExcelReader.ReadExcel(dialog.FileName);
            _excelStatusText.Text = "Loaded " + _excelRows.Count + " rows from Excel.";
            Log("Excel loaded: " + dialog.FileName);
        }
        catch (Exception ex)
        {
            _excelRows = null;
            _excelStatusText.Text = "Excel load failed.";
            Log("Excel load failed: " + ex.Message);
        }
    }
    private async Task ApplyDescriptionsAsync()
    {
        if (_excelRows == null || _excelRows.Count == 0)
        {
            MessageBox.Show("Load an Excel file first.", "ACC Docs", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        ProjectInfo project = _projectCombo.SelectedItem as ProjectInfo;
        if (project == null)
            return;

        int updated = 0;
        int skipped = 0;

        foreach (FileItem file in _fileItems)
        {
            if (file.CanUpdateDescription == false)
            {
                skipped++;
                Log("Skipped (unsupported item type): " + file.DisplayName);
                continue;
            }

            ExcelRow row;
            bool found = _excelRows.TryGetValue(file.DisplayName, out row);
            if (found == false)
            {
                skipped++;
                continue;
            }

            if (string.IsNullOrWhiteSpace(row.Description))
            {
                skipped++;
                continue;
            }

            try
            {
                await _dataClient.UpdateFileDescriptionAsync(project.Id, file.Id, row.Description);
                file.Description = row.Description;
                updated++;
                Log("Updated: " + file.DisplayName);
            }
            catch (Exception ex)
            {
                Log("Failed to update " + file.DisplayName + ": " + ex.Message);
            }
        }

        _statusText.Text = "Update complete";
        Log("Update finished. Updated: " + updated + ", Skipped: " + skipped);
    }

    private void Log(string message)
    {
        if (_logBox == null)
            return;
        _logBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + message + Environment.NewLine);
        _logBox.ScrollToEnd();
    }
}
internal class AccAuthClient
{
    private readonly string _authorizeUrl;
    private readonly string _tokenUrl;
    private readonly string _redirectUri;
    private readonly string _scopes;
    private readonly Action<string> _log;
    private readonly HttpClient _httpClient = new HttpClient();

    public AccAuthClient(string authorizeUrl, string tokenUrl, string redirectUri, string scopes, Action<string> log)
    {
        _authorizeUrl = authorizeUrl;
        _tokenUrl = tokenUrl;
        _redirectUri = redirectUri;
        _scopes = scopes;
        _log = log;
    }

    public async Task<AccAuthSession> AuthenticateAsync(string clientId, string clientSecret)
    {
        string state = Guid.NewGuid().ToString("N");
        string authUrl = BuildAuthUrl(clientId, state);

        _log("Opening browser for login...");

        using (HttpListener listener = new HttpListener())
        {
            listener.Prefixes.Add(_redirectUri);
            listener.Start();

            Process.Start(new ProcessStartInfo(authUrl) { UseShellExecute = true });

            HttpListenerContext context = await listener.GetContextAsync();
            string code = context.Request.QueryString["code"];
            string returnedState = context.Request.QueryString["state"];
            string error = context.Request.QueryString["error"];

            WriteResponse(context.Response, error);

            if (string.IsNullOrEmpty(error) == false)
                throw new InvalidOperationException("Auth error: " + error);

            if (string.IsNullOrEmpty(code))
                throw new InvalidOperationException("Authorization code missing.");

            if (string.Equals(state, returnedState, StringComparison.Ordinal) == false)
                throw new InvalidOperationException("State mismatch. Possible tampering.");

            return await ExchangeCodeAsync(clientId, clientSecret, code);
        }
    }

    public async Task RefreshAsync(string clientId, string clientSecret, AccAuthSession session)
    {
        if (session == null || string.IsNullOrWhiteSpace(session.RefreshToken))
            throw new InvalidOperationException("No refresh token available.");

        Dictionary<string, string> form = new Dictionary<string, string>();
        form["grant_type"] = "refresh_token";
        form["refresh_token"] = session.RefreshToken;

        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenUrl);
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue(
            "Basic",
            Convert.ToBase64String(Encoding.ASCII.GetBytes(clientId + ":" + clientSecret)));
        request.Content = new FormUrlEncodedContent(form);

        HttpResponseMessage response = await _httpClient.SendAsync(request);
        string json = await response.Content.ReadAsStringAsync();
        if (response.IsSuccessStatusCode == false)
            throw new InvalidOperationException("Refresh failed: " + json);

        AccAuthSession updated = AccAuthSession.FromJson(json);
        session.AccessToken = updated.AccessToken;
        session.RefreshToken = updated.RefreshToken;
        session.ExpiresAtUtc = updated.ExpiresAtUtc;
    }

    private string BuildAuthUrl(string clientId, string state)
    {
        string query =
            "response_type=code" +
            "&client_id=" + Uri.EscapeDataString(clientId) +
            "&redirect_uri=" + Uri.EscapeDataString(_redirectUri) +
            "&scope=" + Uri.EscapeDataString(_scopes) +
            "&state=" + Uri.EscapeDataString(state) +
            "&prompt=login";
        return _authorizeUrl + "?" + query;
    }

    private async Task<AccAuthSession> ExchangeCodeAsync(string clientId, string clientSecret, string code)
    {
        Dictionary<string, string> form = new Dictionary<string, string>();
        form["grant_type"] = "authorization_code";
        form["code"] = code;
        form["redirect_uri"] = _redirectUri;

        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenUrl);
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue(
            "Basic",
            Convert.ToBase64String(Encoding.ASCII.GetBytes(clientId + ":" + clientSecret)));
        request.Content = new FormUrlEncodedContent(form);

        HttpResponseMessage response = await _httpClient.SendAsync(request);
        string json = await response.Content.ReadAsStringAsync();
        if (response.IsSuccessStatusCode == false)
            throw new InvalidOperationException("Token exchange failed: " + json);

        return AccAuthSession.FromJson(json);
    }

    private void WriteResponse(HttpListenerResponse response, string error)
    {
        string message = string.IsNullOrEmpty(error)
            ? "You can close this window and return to Revit."
            : "Authentication failed. You can close this window.";

        byte[] buffer = Encoding.UTF8.GetBytes("<html><body><h3>" + message + "</h3></body></html>");
        response.ContentLength64 = buffer.Length;
        using (Stream output = response.OutputStream)
        {
            output.Write(buffer, 0, buffer.Length);
        }
    }
}
internal class AccDataClient
{
    private readonly AccAuthClient _authClient;
    private readonly AccAuthSession _session;
    private readonly Action<string> _log;
    private readonly HttpClient _httpClient = new HttpClient();

    public AccDataClient(AccAuthClient authClient, AccAuthSession session, Action<string> log)
    {
        _authClient = authClient;
        _session = session;
        _log = log;
    }

    public async Task<List<HubInfo>> GetHubsAsync()
    {
        string json = await GetAsync("https://developer.api.autodesk.com/project/v1/hubs");
        Dictionary<string, object> root = JsonHelper.Deserialize(json);
        List<HubInfo> result = new List<HubInfo>();
        foreach (Dictionary<string, object> item in JsonHelper.GetArray(root, "data"))
        {
            string id = JsonHelper.GetString(item, "id");
            Dictionary<string, object> attrs = JsonHelper.GetDict(item, "attributes");
            string name = JsonHelper.GetString(attrs, "name");
            result.Add(new HubInfo { Id = id, Name = name });
        }
        return result;
    }

    public async Task<List<ProjectInfo>> GetProjectsAsync(string hubId)
    {
        string json = await GetAsync("https://developer.api.autodesk.com/project/v1/hubs/" + Uri.EscapeDataString(hubId) + "/projects");
        Dictionary<string, object> root = JsonHelper.Deserialize(json);
        List<ProjectInfo> result = new List<ProjectInfo>();
        foreach (Dictionary<string, object> item in JsonHelper.GetArray(root, "data"))
        {
            string id = JsonHelper.GetString(item, "id");
            Dictionary<string, object> attrs = JsonHelper.GetDict(item, "attributes");
            string name = JsonHelper.GetString(attrs, "name");
            result.Add(new ProjectInfo { Id = id, Name = name });
        }
        return result;
    }

    public async Task<List<FolderNode>> GetTopFoldersAsync(string hubId, string projectId)
    {
        string url = "https://developer.api.autodesk.com/project/v1/hubs/" + Uri.EscapeDataString(hubId) + "/projects/" + Uri.EscapeDataString(projectId) + "/topFolders";
        string json = await GetAsync(url);
        Dictionary<string, object> root = JsonHelper.Deserialize(json);
        List<FolderNode> result = new List<FolderNode>();

        foreach (Dictionary<string, object> item in JsonHelper.GetArray(root, "data"))
        {
            string id = JsonHelper.GetString(item, "id");
            Dictionary<string, object> attrs = JsonHelper.GetDict(item, "attributes");
            string name = JsonHelper.GetString(attrs, "displayName");
            result.Add(new FolderNode(id, name));
        }

        return result;
    }
    public async Task<List<FolderNode>> GetFolderChildrenAsync(string projectId, string folderId)
    {
        List<FolderNode> children = new List<FolderNode>();
        List<FileItem> files = new List<FileItem>();

        string url = "https://developer.api.autodesk.com/data/v1/projects/" + Uri.EscapeDataString(projectId) + "/folders/" + Uri.EscapeDataString(folderId) + "/contents";
        string json = await GetAsync(url);

        Dictionary<string, object> root = JsonHelper.Deserialize(json);
        ParseFolderContents(root, children, files);

        return children;
    }

    public async Task<List<FileItem>> GetFilesInFolderAsync(string projectId, string folderId)
    {
        List<FolderNode> folders = new List<FolderNode>();
        List<FileItem> files = new List<FileItem>();

        string url = "https://developer.api.autodesk.com/data/v1/projects/" + Uri.EscapeDataString(projectId) + "/folders/" + Uri.EscapeDataString(folderId) + "/contents";

        string nextUrl = url;
        while (!string.IsNullOrEmpty(nextUrl))
        {
            string json = await GetAsync(nextUrl);
            Dictionary<string, object> root = JsonHelper.Deserialize(json);
            ParseFolderContents(root, folders, files);
            nextUrl = JsonHelper.GetNextLink(root);
        }

        return files;
    }

    public async Task UpdateFileDescriptionAsync(string projectId, string itemId, string description)
    {
        string url = "https://developer.api.autodesk.com/data/v1/projects/" + Uri.EscapeDataString(projectId) + "/items/" + Uri.EscapeDataString(itemId);

        string body = "{\"jsonapi\":{\"version\":\"1.0\"},\"data\":{\"type\":\"items\",\"id\":\"" + EscapeJson(itemId) + "\",\"attributes\":{\"description\":\"" + EscapeJson(description) + "\"}}}";

        HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), url);
        request.Content = new StringContent(body, Encoding.UTF8, "application/vnd.api+json");

        HttpResponseMessage response = await SendAsync(request);
        if (response.IsSuccessStatusCode == false)
        {
            string error = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException("Update failed: " + error);
        }
    }
    private async Task<string> GetAsync(string url)
    {
        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
        HttpResponseMessage response = await SendAsync(request);
        string json = await response.Content.ReadAsStringAsync();
        if (response.IsSuccessStatusCode == false)
            throw new InvalidOperationException(json);
        return json;
    }

    private async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request)
    {
        await EnsureTokenAsync();
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _session.AccessToken);
        request.Headers.Accept.Clear();
        request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
        request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/vnd.api+json"));

        return await _httpClient.SendAsync(request);
    }

    private async Task EnsureTokenAsync()
    {
        if (_session == null)
            throw new InvalidOperationException("No auth session.");

        if (_session.ExpiresAtUtc <= DateTime.UtcNow.AddMinutes(2))
        {
            _log("Refreshing token...");
            await _authClient.RefreshAsync(_session.ClientId, _session.ClientSecret, _session);
        }
    }

    private void ParseFolderContents(Dictionary<string, object> root, List<FolderNode> folders, List<FileItem> files)
    {
        foreach (Dictionary<string, object> item in JsonHelper.GetArray(root, "data"))
        {
            string type = JsonHelper.GetString(item, "type");
            Dictionary<string, object> attrs = JsonHelper.GetDict(item, "attributes");

            if (string.Equals(type, "folders", StringComparison.OrdinalIgnoreCase))
            {
                string id = JsonHelper.GetString(item, "id");
                string name = JsonHelper.GetString(attrs, "displayName");
                folders.Add(new FolderNode(id, name));
            }
            else if (string.Equals(type, "items", StringComparison.OrdinalIgnoreCase))
            {
                string id = JsonHelper.GetString(item, "id");
                string name = JsonHelper.GetString(attrs, "displayName");
                string desc = JsonHelper.GetString(attrs, "description");
                Dictionary<string, object> ext = JsonHelper.GetDict(attrs, "extension");
                string extType = JsonHelper.GetString(ext, "type");
                bool canUpdate = string.Equals(extType, "items:autodesk.bim360:File", StringComparison.OrdinalIgnoreCase);
                files.Add(new FileItem
                {
                    Id = id,
                    DisplayName = name,
                    Description = desc,
                    ExtensionType = extType,
                    CanUpdateDescription = canUpdate
                });
            }
        }
    }

    private static string EscapeJson(string value)
    {
        if (value == null)
            return string.Empty;
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n");
    }
}
internal class AccAuthSession
{
    public string AccessToken { get; set; }
    public string RefreshToken { get; set; }
    public DateTime ExpiresAtUtc { get; set; }
    public string ClientId { get; set; }
    public string ClientSecret { get; set; }

    public static AccAuthSession FromJson(string json)
    {
        Dictionary<string, object> root = JsonHelper.Deserialize(json);
        string accessToken = JsonHelper.GetString(root, "access_token");
        string refreshToken = JsonHelper.GetString(root, "refresh_token");
        int expiresIn = JsonHelper.GetInt(root, "expires_in");

        return new AccAuthSession
        {
            AccessToken = accessToken,
            RefreshToken = refreshToken,
            ExpiresAtUtc = DateTime.UtcNow.AddSeconds(expiresIn > 0 ? expiresIn : 3000)
        };
    }
}
internal static class JsonHelper
{
    private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer();

    static JsonHelper()
    {
        Serializer.MaxJsonLength = int.MaxValue;
    }

    public static Dictionary<string, object> Deserialize(string json)
    {
        return Serializer.Deserialize<Dictionary<string, object>>(json);
    }

    public static List<Dictionary<string, object>> GetArray(Dictionary<string, object> dict, string key)
    {
        List<Dictionary<string, object>> result = new List<Dictionary<string, object>>();
        if (dict == null || dict.ContainsKey(key) == false || dict[key] == null)
            return result;

        IEnumerable items = dict[key] as IEnumerable;
        if (items == null)
            return result;

        foreach (object obj in items)
        {
            Dictionary<string, object> item = obj as Dictionary<string, object>;
            if (item != null)
                result.Add(item);
        }

        return result;
    }

    public static Dictionary<string, object> GetDict(Dictionary<string, object> dict, string key)
    {
        if (dict == null || dict.ContainsKey(key) == false)
            return new Dictionary<string, object>();

        Dictionary<string, object> value = dict[key] as Dictionary<string, object>;
        return value ?? new Dictionary<string, object>();
    }

    public static string GetString(Dictionary<string, object> dict, string key)
    {
        if (dict == null || dict.ContainsKey(key) == false || dict[key] == null)
            return string.Empty;
        return dict[key].ToString();
    }

    public static int GetInt(Dictionary<string, object> dict, string key)
    {
        if (dict == null || dict.ContainsKey(key) == false || dict[key] == null)
            return 0;

        int value;
        if (int.TryParse(dict[key].ToString(), out value))
            return value;

        return 0;
    }

    public static string GetNextLink(Dictionary<string, object> dict)
    {
        Dictionary<string, object> links = GetDict(dict, "links");
        if (links.Count == 0)
            return null;

        if (links.ContainsKey("next") == false)
            return null;

        Dictionary<string, object> next = links["next"] as Dictionary<string, object>;
        if (next == null)
            return null;

        return GetString(next, "href");
    }
}
internal class HubInfo
{
    public string Id { get; set; }
    public string Name { get; set; }
    public override string ToString() { return Name; }
}

internal class ProjectInfo
{
    public string Id { get; set; }
    public string Name { get; set; }
    public override string ToString() { return Name; }
}

internal class FolderNode
{
    public string Id { get; private set; }
    public string Name { get; private set; }
    public ObservableCollection<FolderNode> Children { get; private set; }
    public bool IsPlaceholder { get; private set; }
    public bool IsLoaded { get; set; }
    public bool IsLoading { get; set; }

    public FolderNode(string id, string name, bool isPlaceholder = false)
    {
        Id = id;
        Name = name;
        IsPlaceholder = isPlaceholder;
        Children = new ObservableCollection<FolderNode>();
        if (isPlaceholder == false)
            Children.Add(new FolderNode(string.Empty, string.Empty, true));
    }
}

internal class FileItem : INotifyPropertyChanged
{
    private string _description;

    public string Id { get; set; }
    public string DisplayName { get; set; }
    public string ExtensionType { get; set; }
    public bool CanUpdateDescription { get; set; }
    public string Description
    {
        get { return _description; }
        set
        {
            if (_description == value)
                return;
            _description = value;
            OnPropertyChanged("Description");
        }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    private void OnPropertyChanged(string propertyName)
    {
        if (PropertyChanged != null)
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
    }
}

internal class ExcelRow
{
    public string FileName { get; set; }
    public string Description { get; set; }
}
internal static class ExcelReader
{
    public static Dictionary<string, ExcelRow> ReadExcel(string path)
    {
        Dictionary<string, ExcelRow> result = new Dictionary<string, ExcelRow>(StringComparer.OrdinalIgnoreCase);

        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
            throw new InvalidOperationException("Excel is not installed.");

        dynamic excel = Activator.CreateInstance(excelType);
        excel.Visible = false;
        dynamic workbook = null;
        dynamic sheet = null;
        dynamic usedRange = null;

        try
        {
            workbook = excel.Workbooks.Open(path, ReadOnly: true);
            sheet = workbook.Worksheets[1];
            usedRange = sheet.UsedRange;
            object[,] values = usedRange.Value2 as object[,];

            if (values == null)
                return result;

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            Dictionary<string, int> headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int col = 1; col <= cols; col++)
            {
                object headerVal = values[1, col];
                if (headerVal == null)
                    continue;
                string header = headerVal.ToString().Trim();
                if (string.IsNullOrWhiteSpace(header))
                    continue;
                if (headers.ContainsKey(header) == false)
                    headers.Add(header, col);
            }

            int descriptionCol = headers.ContainsKey("description") ? headers["description"] : -1;

            for (int row = 2; row <= rows; row++)
            {
                object fileVal = values[row, 1];
                if (fileVal == null)
                    continue;
                string fileName = fileVal.ToString().Trim();
                if (string.IsNullOrWhiteSpace(fileName))
                    continue;

                string description = string.Empty;
                if (descriptionCol > 0)
                {
                    object descVal = values[row, descriptionCol];
                    if (descVal != null)
                        description = descVal.ToString().Trim();
                }

                ExcelRow rowData = new ExcelRow
                {
                    FileName = fileName,
                    Description = description
                };

                result[fileName] = rowData;
            }

            return result;
        }
        finally
        {
            if (workbook != null)
                workbook.Close(false);
            if (excel != null)
                excel.Quit();

            ReleaseComObject(usedRange);
            ReleaseComObject(sheet);
            ReleaseComObject(workbook);
            ReleaseComObject(excel);
        }
    }

    private static void ReleaseComObject(object obj)
    {
        try
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.FinalReleaseComObject(obj);
        }
        catch { }
    }
}
