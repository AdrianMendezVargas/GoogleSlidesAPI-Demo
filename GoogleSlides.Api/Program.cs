using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Util.Store;
using GoogleSlides.Api.Data;
using GoogleSlides.Api.Models.Domain;
using GoogleSlides.Api.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using System.Net;

var builder = WebApplication.CreateBuilder(args);
string[] Scopes = { SlidesService.Scope.Presentations, DriveService.Scope.Drive };


// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddDbContext<ApplicationDbContext>(options =>
{
    options.UseSqlite(builder.Configuration["ConnectionStrings:Sqlite"]);
});

builder.Services.AddScoped<AuthConfigService>();

//builder.Services.AddSingleton<ClientSecrets>(GetClientSecrets());

//var credentials = GetCredentials();

var credentials = GoogleCredential.FromJson(File.ReadAllText("service-account_veloci.json"))
    .CreateScoped(SlidesService.Scope.Presentations, DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets);

//builder.Services.AddSingleton<UserCredential>(credentials);

//using (var provider = builder.Services.BuildServiceProvider())
//{
//    var authConfigService = provider.GetRequiredService<AuthConfigService>();
//    var context = provider.GetRequiredService<ApplicationDbContext>();


//    authConfigService.Save(new AuthConfig()
//    {
//        Token = credentials..Token.AccessToken,
//        RefreshToken = credentials.Token.RefreshToken,
//        DurationInSeconds = (int)credentials.Token.ExpiresInSeconds,
//        IssueTime = credentials.Token.IssuedUtc
//    });
//}


builder.Services.AddSingleton(sp =>
{
    return new DriveService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credentials,
        ApplicationName = builder.Configuration["GoogleAPI:AppName"]
    });
});

builder.Services.AddSingleton(sp =>
{
    return new SlidesService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credentials,
        ApplicationName = builder.Configuration["GoogleAPI:AppName"]
    });
});

builder.Services.AddSingleton(sp =>
{
    return new SheetsService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credentials,
        ApplicationName = builder.Configuration["GoogleAPI:AppName"]
    });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();

ClientSecrets GetClientSecrets()
{

    return new ClientSecrets();

    ClientSecrets clientSecrets;

    using (var stream = new FileStream("service-account_veloci.json", FileMode.Open, FileAccess.Read))
    {
        clientSecrets = GoogleClientSecrets.FromStream(stream).Secrets;
    }
    return clientSecrets;
}

UserCredential GetCredentials()
{
    UserCredential credential;

    using (var stream = new FileStream("service-account_veloci.json", FileMode.Open, FileAccess.Read))
    {
        string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

        credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");

        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            GoogleClientSecrets.FromStream(stream).Secrets,
            Scopes,
            "user",
            CancellationToken.None,
            new FileDataStore(credPath, true)).Result;
        // Console.WriteLine("Credential file saved to: " + credPath);
    }

    return credential;
}