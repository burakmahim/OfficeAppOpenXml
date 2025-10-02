

var builder = WebApplication.CreateBuilder(args);



// 2) Servisler
builder.Services.AddControllersWithViews();

var app = builder.Build();

// pipeline...
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseRouting();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Presentation}/{action=Index}/{id?}");

app.Run();
