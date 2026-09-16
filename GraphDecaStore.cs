using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;

public sealed record DecaItem(string? Token, string DriveId, string ItemId, string FileName);

/// <summary>
/// Busca y descarga los PDF de albaranes guardados en una biblioteca de SharePoint mediante Graph.
/// Configuración (App Settings de la Function): TENANT_ID, CLIENT_ID, CLIENT_SECRET, SP_SITE_ID, SP_LIST_ID
/// </summary>
public sealed class GraphDecaStore
{
    private static readonly string[] Scopes = ["https://graph.microsoft.com/.default"];
    private readonly HttpClient _http;
    private readonly TokenCredential _cred;
    private readonly string _siteId;
    private readonly string _listId;

    public GraphDecaStore(IHttpClientFactory factory)
    {
        _http = factory.CreateClient();
        _cred = new ClientSecretCredential(Env("TENANT_ID"), Env("CLIENT_ID"), Env("CLIENT_SECRET"));
        _siteId = Env("SP_SITE_ID");
        _listId = Env("SP_LIST_ID");
    }

    public async Task<DecaItem?> FindAsync(string albaran, CancellationToken ct)
    {
        // albaran ya viene validado (solo letras, números, guiones), así que no hay inyección OData
        var filter = Uri.EscapeDataString($"fields/Albaran eq '{albaran}'");
        var url = $"https://graph.microsoft.com/v1.0/sites/{_siteId}/lists/{_listId}/items" +
                  $"?$filter={filter}&$expand=fields($select=Albaran,Token),driveItem&$top=2";

        using var req = await AuthorizedAsync(url, ct);
        req.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
        using var resp = await _http.SendAsync(req, ct);
        resp.EnsureSuccessStatusCode();

        using var doc = await JsonDocument.ParseAsync(await resp.Content.ReadAsStreamAsync(ct), cancellationToken: ct);
        var items = doc.RootElement.GetProperty("value");
        if (items.GetArrayLength() != 1) return null; // 0 = no existe, 2+ = duplicado (no servimos nada)

        var it = items[0];
        var fields = it.GetProperty("fields");
        var drive = it.GetProperty("driveItem");

        string? token = fields.TryGetProperty("Token", out var t) ? t.GetString() : null;

        return new DecaItem(
            token,
            drive.GetProperty("parentReference").GetProperty("driveId").GetString()!,
            drive.GetProperty("id").GetString()!,
            drive.GetProperty("name").GetString()!);
    }

    /// <summary>Descarga el contenido. Graph responde con 302 a una URL pre-autenticada; HttpClient la sigue.</summary>
    public async Task<HttpResponseMessage> DownloadAsync(DecaItem item, CancellationToken ct)
    {
        var url = $"https://graph.microsoft.com/v1.0/drives/{item.DriveId}/items/{item.ItemId}/content";
        var req = await AuthorizedAsync(url, ct);
        return await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
    }

    private async Task<HttpRequestMessage> AuthorizedAsync(string url, CancellationToken ct)
    {
        var token = await _cred.GetTokenAsync(new TokenRequestContext(Scopes), ct);
        var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        return req;
    }

    private static string Env(string name) =>
        Environment.GetEnvironmentVariable(name) ?? throw new InvalidOperationException($"Falta la variable {name}");
}
