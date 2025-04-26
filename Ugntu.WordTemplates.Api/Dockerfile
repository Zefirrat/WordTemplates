FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
USER $APP_UID
WORKDIR /app
EXPOSE 8080
EXPOSE 8081

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["Ugntu.WordTemplates.Api/Ugntu.WordTemplates.Api.csproj", "Ugntu.WordTemplates.Api/"]
RUN dotnet restore "Ugntu.WordTemplates.Api/Ugntu.WordTemplates.Api.csproj"
COPY . .
WORKDIR "/src/Ugntu.WordTemplates.Api"
RUN dotnet build "./Ugntu.WordTemplates.Api.csproj" -c $BUILD_CONFIGURATION -o /app/build

FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./Ugntu.WordTemplates.Api.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Ugntu.WordTemplates.Api.dll"]
