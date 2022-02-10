#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.
FROM mcr.microsoft.com/dotnet/aspnet:6.0.2-alpine3.14-amd64 AS base
#RUN apk add libgdiplus --update-cache --repository http://dl-3.alpinelinux.org/alpine/edge/testing/ --allow-untrusted
#RUN apk --no-cache add msttcorefonts-installer fontconfig freetype-dev libjpeg-turbo-dev libpng-dev && \
    #update-ms-fonts && \
    #fc-cache -f

WORKDIR /app

EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/aspnet:6.0.2-alpine3.14-amd64 AS build
WORKDIR /src
COPY ["EPPlus.DockerSamples.Alpine.csproj", ""]
RUN dotnet restore "./EPPlus.DockerSamples.Alpine.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "EPPlus.DockerSamples.Alpine.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "EPPlus.DockerSamples.Alpine.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "EPPlus.DockerSamples.Alpine.dll"]
