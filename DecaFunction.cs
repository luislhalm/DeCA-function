using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

/// <summary>
/// GET /d/{albaran}/{token}  ->  descarga directa del PDF (sin login, sin página intermedia),
/// como exige la Resolución de 5 de junio de 2026 para los DeCA.
/// </summary>
public sealed partial class DecaFunction(GraphDecaStore store, ILogger<DecaFunction> log)
{
    [GeneratedRegex("^[A-Za-z0-9_-]{1,60}$")]
    private static partial Regex SafeId();

    [Function("Deca")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "d/{albaran}/{token}")] HttpRequestData req,
        string albaran,
        string token,
        CancellationToken ct)
    {
        if (!SafeId().IsMatch(albaran) || !SafeId().IsMatch(token))
            return req.CreateResponse(HttpStatusCode.NotFound);

        var item = await store.FindAsync(albaran, ct);
        if (item?.Token is null || !SameToken(item.Token, token))
        {
            log.LogWarning("DeCA no encontrado o token inválido: {Albaran}", albaran);
            return req.CreateResponse(HttpStatusCode.NotFound);
        }

        // Sin caducidad: el enlace funciona mientras el fichero exista en SharePoint.
        // (La Resolución permite, pero no obliga, a desactivarlo 7 días tras el fin del servicio.)

        using var file = await store.DownloadAsync(item, ct);
        if (!file.IsSuccessStatusCode)
        {
            log.LogError("Error {Status} descargando {Albaran} de SharePoint", (int)file.StatusCode, albaran);
            return req.CreateResponse(HttpStatusCode.BadGateway);
        }

        var resp = req.CreateResponse(HttpStatusCode.OK);
        resp.Headers.Add("Content-Type", "application/pdf");
        resp.Headers.Add("Content-Disposition", $"attachment; filename=\"{item.FileName}\"");
        resp.Headers.Add("Cache-Control", "no-store");
        resp.Headers.Add("X-Robots-Tag", "noindex");
        await file.Content.CopyToAsync(resp.Body, ct);

        // Queda registro en Application Insights de cada descarga (útil ante una inspección)
        log.LogInformation("DeCA servido: {Albaran}", albaran);
        return resp;
    }

    private static bool SameToken(string a, string b) =>
        CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(a), Encoding.UTF8.GetBytes(b));
}
