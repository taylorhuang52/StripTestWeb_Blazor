using System.Text;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.Extensions.DependencyInjection;
using StripTestBlazor.Services;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<StripTestBlazor.App>("#app");
builder.Services.AddScoped<ReportService>();

await builder.Build().RunAsync();
