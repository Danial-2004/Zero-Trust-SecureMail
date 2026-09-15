using System.IdentityModel.Tokens.Jwt;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Logging;
using Microsoft.IdentityModel.Tokens;
using Microsoft.OpenApi.Models;
using SecureMailPolicyApi.Services;

var builder = WebApplication.CreateBuilder(args);

// Show detailed identity errors during development
IdentityModelEventSource.ShowPII = true;

builder.Logging.ClearProviders();
builder.Logging.AddConsole();
builder.Logging.AddDebug();
builder.Logging.SetMinimumLevel(LogLevel.Debug);

builder.Services.AddControllers();

builder.Services
    .AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.Authority = "http://localhost:8080/realms/SecureMailRealm";
        options.MetadataAddress = "http://localhost:8080/realms/SecureMailRealm/.well-known/openid-configuration";
        options.RequireHttpsMetadata = false;
        options.MapInboundClaims = false;
        options.SaveToken = true;

        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = false,
            ValidateAudience = false,
            ValidateLifetime = true,
            ClockSkew = TimeSpan.FromMinutes(2),
            NameClaimType = "preferred_username",
            RoleClaimType = "roles"
        };

        options.Events = new JwtBearerEvents
        {
            OnMessageReceived = context =>
            {
                var authHeader = context.Request.Headers["Authorization"].ToString();
                Console.WriteLine("=== OnMessageReceived ===");
                Console.WriteLine("Authorization header present: " + !string.IsNullOrWhiteSpace(authHeader));
                return Task.CompletedTask;
            },
            OnTokenValidated = context =>
            {
                Console.WriteLine("=== TOKEN VALIDATED SUCCESSFULLY ===");
                var sub = context.Principal?.FindFirst("sub")?.Value;
                var user = context.Principal?.FindFirst("preferred_username")?.Value;
                Console.WriteLine("sub: " + sub);
                Console.WriteLine("preferred_username: " + user);
                return Task.CompletedTask;
            },
            OnAuthenticationFailed = context =>
            {
                Console.WriteLine("=== JWT AUTH FAILED ===");
                Console.WriteLine(context.Exception.ToString());
                return Task.CompletedTask;
            },
            OnChallenge = context =>
            {
                Console.WriteLine("=== JWT CHALLENGE ===");
                Console.WriteLine("Error: " + context.Error);
                Console.WriteLine("Description: " + context.ErrorDescription);
                return Task.CompletedTask;
            }
        };
    });

builder.Services.AddAuthorization();
builder.Services.AddSingleton<FilePolicyService>();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "SecureMailPolicyApi",
        Version = "v1"
    });

    c.AddSecurityDefinition("oauth2", new OpenApiSecurityScheme
    {
        Type = SecuritySchemeType.OAuth2,
        Flows = new OpenApiOAuthFlows
        {
            AuthorizationCode = new OpenApiOAuthFlow
            {
                AuthorizationUrl = new Uri("http://localhost:8080/realms/SecureMailRealm/protocol/openid-connect/auth"),
                TokenUrl = new Uri("http://localhost:8080/realms/SecureMailRealm/protocol/openid-connect/token"),
                Scopes = new Dictionary<string, string>
                {
                    { "openid", "OpenID Connect scope" },
                    { "profile", "Profile scope" },
                    { "email", "Email scope" }
                }
            }
        }
    });

    c.AddSecurityRequirement(new OpenApiSecurityRequirement
    {
        {
            new OpenApiSecurityScheme
            {
                Reference = new OpenApiReference
                {
                    Type = ReferenceType.SecurityScheme,
                    Id = "oauth2"
                }
            },
            new[] { "openid", "profile", "email" }
        }
    });
});

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "SecureMailPolicyApi v1");
    c.OAuthClientId("securemail-swagger");
    c.OAuthAppName("SecureMail Policy Swagger");
    c.OAuthUsePkce();
    c.OAuthScopes("openid", "profile", "email");
});

app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

app.Run();