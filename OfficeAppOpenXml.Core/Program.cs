using System.Text;
using Syncfusion.Licensing;
using OfficeAppOpenXmlLibrary.Services;

SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1JGaF5cXGpCf1FpRmJGdld5fUVHYVZUTXxaS00DNHVRdkdmWH1ccnRQRGBfV0BxXEtWYEs=");


var builder = WebApplication.CreateBuilder(args);
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

builder.Services.AddControllersWithViews();

// 1) Connection string'i oku
string cs = builder.Configuration.GetConnectionString("OfficeFilesDb");

// 2) FileService'i baðlantý ile ekle
builder.Services.AddSingleton<FileService>(new FileService(cs));

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Presentation}/{action=Index}/{id?}");

app.Run();
