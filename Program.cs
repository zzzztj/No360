using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using No360.Configuration;
using No360.Services;

namespace No360;

public static class Program
{
    public static async Task Main(string[] args)
    {
        var host = Host.CreateDefaultBuilder(args)
            .UseWindowsService()
            .ConfigureAppConfiguration(cfg =>
            {
                cfg.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            })
            .ConfigureServices((ctx, services) =>
            {
                services.AddSingleton(
                    ctx.Configuration.GetSection("BlockRules").Get<BlockRules>() ?? new BlockRules());
                services.AddSingleton(new RefereeSettings(ctx.Configuration));
                services.AddHostedService<RefereeService>();
            })
            .Build();

        await host.RunAsync().ConfigureAwait(false);
    }
}
