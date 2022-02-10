#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.
FROM mcr.microsoft.com/dotnet/aspnet:6.0.2-nanoserver-20H2 AS base

WORKDIR /app

EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/aspnet:6.0.2-nanoserver-20H2 AS build
WORKDIR /src
COPY ["EPPlus.DockerSamples.Windows.Nano.Server.csproj", ""]
RUN dotnet restore "./EPPlus.DockerSamples.Windows.Nano.Server.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "EPPlus.DockerSamples.Windows.Nano.Server.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "EPPlus.DockerSamples.Windows.Nano.Server.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "EPPlus.DockerSamples.Windows.Nano.Server.dll"]
