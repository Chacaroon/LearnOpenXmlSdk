FROM mcr.microsoft.com/dotnet/runtime:6.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["LearnOpenXml/LearnOpenXml.csproj", "LearnOpenXml/"]
RUN dotnet restore "LearnOpenXml/LearnOpenXml.csproj"
COPY . .
WORKDIR "/src/LearnOpenXml"
RUN dotnet build "LearnOpenXml.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "LearnOpenXml.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "LearnOpenXml.dll"]
