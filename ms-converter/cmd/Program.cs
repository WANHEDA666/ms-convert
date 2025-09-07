using ms_converter.service;

var builder = Host.CreateApplicationBuilder(args);
builder.Services.AddWindowsService(options => options.ServiceName = "MsConverter");

builder.Logging.ClearProviders();
builder.Logging.AddSimpleConsole(o =>
{
    o.TimestampFormat = "HH:mm:ss ";
    o.SingleLine = true;
});

builder.Configuration.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
builder.Services.Configure<RabbitOptions>(builder.Configuration.GetSection("Rabbit"));
builder.Services.Configure<StorageOptions>(builder.Configuration.GetSection("Storage"));
builder.Services.Configure<S3Options>(builder.Configuration.GetSection("S3"));

builder.Services.AddHttpClient<Storage>();
builder.Services.AddTransient<Converter>();
builder.Services.AddSingleton<S3Uploader>();
builder.Services.AddHostedService<Consumer>();

var app = builder.Build();
await app.RunAsync();