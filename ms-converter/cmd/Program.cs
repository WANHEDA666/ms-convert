using ms_converter.service;

var builder = Host.CreateApplicationBuilder(args);
builder.Services.Configure<RabbitOptions>(builder.Configuration.GetSection("Rabbit"));
builder.Services.Configure<StorageOptions>(builder.Configuration.GetSection("Storage"));

builder.Services.AddHttpClient<Storage>();
builder.Services.AddTransient<Converter>();
builder.Services.AddHostedService<Consumer>();

var host = builder.Build();
host.Run();