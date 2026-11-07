# syntax=docker/dockerfile:1

# ---------------------------------------------------------------------------
# Build stage
# ---------------------------------------------------------------------------
# The build runs on the builder's own platform: the published app is portable (no native host), so the arm64 image needs no emulation here.
FROM --platform=$BUILDPLATFORM mcr.microsoft.com/dotnet/sdk:10.0 AS build
ARG Channel=local
ARG BuildNumber=0
ARG BuildDate=
WORKDIR /src

# Restore first (better layer caching).
COPY ["global.json", "Directory.Build.props", "./"]
COPY ["src/MatOffice/MatOffice.csproj", "src/MatOffice/"]
RUN dotnet restore "src/MatOffice/MatOffice.csproj"

COPY src src
# Fall back to today's date when BuildDate is not supplied, so the version never ends in "-".
RUN BUILD_DATE="${BuildDate:-$(date -u +%Y%m%d)}" && \
    dotnet publish "src/MatOffice/MatOffice.csproj" \
    -c Release -o /app/publish --no-restore \
    -p:UseAppHost=false \
    -p:Channel=$Channel -p:BuildNumber=$BuildNumber -p:BuildDate=$BUILD_DATE

# ---------------------------------------------------------------------------
# Runtime stage (ASP.NET Core base image)
# ---------------------------------------------------------------------------
FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime

LABEL org.opencontainers.image.title="MatOffice" \
      org.opencontainers.image.description="Office im Browser – Texte, Tabellen, Präsentationen und Datenbanken in einer App, mit Word-, Excel- und PowerPoint-Dateien" \
      org.opencontainers.image.source="https://github.com/Real-TTX/MatOffice"

WORKDIR /app

ENV ASPNETCORE_URLS=http://+:4820 \
    HTTP_PORTS="" \
    ASPNETCORE_ENVIRONMENT=Production \
    DOTNET_gcServer=0 \
    TZ=Europe/Berlin

EXPOSE 4820

COPY --from=build /app/publish .

# One volume for what the app keeps in files: the keys that protect the sessions. The database is the other container.
RUN mkdir -p /data/keys && chown -R $APP_UID:$APP_UID /data
VOLUME /data
USER $APP_UID

# The app checks itself on /api/health (200 for ok and degraded, 503 for down), so the image needs no curl.
HEALTHCHECK --interval=30s --timeout=8s --start-period=30s --retries=3 \
    CMD ["dotnet", "MatOffice.dll", "--healthcheck"]

ENTRYPOINT ["dotnet", "MatOffice.dll"]
