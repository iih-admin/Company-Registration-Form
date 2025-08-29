# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

# Copy solution and restore dependencies
COPY *.sln .
COPY CompanyRegistrationForm/*.csproj ./CompanyRegistrationForm/
RUN dotnet restore

# Copy everything else and publish
COPY CompanyRegistrationForm/. ./CompanyRegistrationForm/
WORKDIR /app/CompanyRegistrationForm
RUN dotnet publish -c Release -o out

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app/CompanyRegistrationForm/out .
ENTRYPOINT ["dotnet", "CompanyRegistrationForm.dll"]
