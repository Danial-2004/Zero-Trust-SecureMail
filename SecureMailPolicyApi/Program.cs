using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Logging;
using Microsoft.IdentityModel.Tokens;
using Microsoft.OpenApi.Models;
using SecureMailPolicyApi.Security;
using SecureMailPolicyApi.Services;

var builder = WebApplication.CreateBuilder(args);

// Never emit personally identifiable identity-model diagnostics by default.
IdentityModelEventSource.ShowPII = false;

builder.Logging.ClearProviders();
builder.Logging.AddConsole();
builder.Logging.SetMinimumLevel(LogLevel.Information);

builder.Services.AddControllers();

string issuer = builder.Configuration["Jwt:Issuer"] ?? string.Empty;
string[]? audiences = builder.Configuration.GetSection("Jwt:Audiences").Get<string[]>();
string metadataAddress = builder.Configuration["Jwt:MetadataAddress"] ?? string.Empty;
string[]? allowedAlgorithms = builder.Configuration.GetSection("Jwt:AllowedAlgorithms").Get<string[]>();

if (string.IsNullOrWhiteSpace(issuer))
    throw new InvalidOperationException("Jwt:Issuer is required. Refusing to start with issuer validation disabled.");

if (audiences == null || audiences.Length == 0)
    throw new InvalidOperationException("Jwt:Audiences is required. Refusing to start with audience validation disabled.");

if (allowedAlgorithms == null || allowedAlgorithms.Length == 0)
{
    allowedAlgorithms = new[] { "RS256", "RS384", "RS512", "PS256", "PS384", "PS512" };
}

if (string.IsNullOrWhiteSpace(metadataAddress))
    metadataAddress = issuer.TrimEnd('/') + "/.well-known/openid-configuration";

builder.Services
    .AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.Authority = issuer;
        options.MetadataAddress = metadataAddress;
        // HTTP metadata is permitted only for local development with localhost.
        options.RequireHttpsMetadata = !builder.Environment.IsDevelopment();
        options.MapInboundClaims = false;
        options.SaveToken = false;

        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidIssuer = issuer,
            ValidateAudience = true,
            ValidAudiences = audiences,
            ValidateIssuerSigningKey = true,
            ValidateLifetime = true,
            RequireExpirationTime = true,
            RequireSignedTokens = true,
            ClockSkew = TimeSpan.FromSeconds(30),
            ValidAlgorithms = allowedAlgorithms,
            NameClaimType = "preferred_username",
            RoleClaimType = "roles"
        };
    });

builder.Services.AddAuthorization(options =>
{
    options.AddPolicy("AdminOnly", policy =>
        policy.RequireAssertion(context => KeycloakRoles.HasRealmRole(context.User, "admin")));
});

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
                AuthorizationUrl = new Uri(issuer.TrimEnd('/') + "/protocol/openid-connect/auth"),
                TokenUrl = new Uri(issuer.TrimEnd('/') + "/protocol/openid-connect/token"),
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
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "SecureMail Policy API v1");
    c.OAuthClientId("securemail-swagger");
    c.OAuthAppName("SecureMail Policy Swagger");
    c.OAuthUsePkce();
    c.OAuthScopes("openid", "profile", "email");
});

app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();
app.Run();
