using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Cors;
using System.Web.Http.Cors;

public class CorsHandle : Attribute, ICorsPolicyProvider
{
    private readonly CorsPolicy _policy;

    public CorsHandle()
    {
        // 建立一個跨網域存取的原則物件
        _policy = new CorsPolicy
        {
            AllowAnyMethod = true,
            AllowAnyHeader = true
        };

        // 在這裡透過資料庫或是設定的方式，可動態加入允許存取的來源網域清單
        // Add allowed origins.
        _policy.Origins.Add("http://myclient.azurewebsites.net");
        _policy.Origins.Add("http://www.contoso.com");
        _policy.Origins.Add("*");
    }

    public Task<CorsPolicy> GetCorsPolicyAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        return Task.FromResult(_policy);
    }
}