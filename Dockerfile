# ---- build stage ----
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

COPY *.sln ./
COPY CustomerInfo/CustomerInfo.csproj CustomerInfo/
COPY global.json ./

RUN dotnet restore CustomerInfo/CustomerInfo.csproj

COPY . .
RUN dotnet publish CustomerInfo/CustomerInfo.csproj -c Release -o /app/publish /p:UseAppHost=false

# ---- runtime stage ----
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app
COPY --from=build /app/publish .

EXPOSE 8080
ENTRYPOINT ["dotnet", "CustomerInfo.dll"]
