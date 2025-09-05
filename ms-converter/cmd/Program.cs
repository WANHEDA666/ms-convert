using ms_converter.service;

var builder = Host.CreateApplicationBuilder(args);
builder.Configuration.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
builder.Services.Configure<RabbitOptions>(builder.Configuration.GetSection("Rabbit"));
builder.Services.Configure<StorageOptions>(builder.Configuration.GetSection("Storage"));
builder.Services.Configure<S3Options>(builder.Configuration.GetSection("S3"));

builder.Services.AddHttpClient<Storage>();
builder.Services.AddTransient<Converter>();
builder.Services.AddTransient<S3Uploader>(); 
builder.Services.AddHostedService<Consumer>();

var app = builder.Build();
app.Run();