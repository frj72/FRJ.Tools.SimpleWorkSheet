FROM mcr.microsoft.com/dotnet/runtime:10.0

ENV DEBIAN_FRONTEND=noninteractive

RUN echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections

RUN apt-get update && apt-get install -y \
    fontconfig \
    fonts-liberation \
    fonts-dejavu-core \
    libfontconfig1 \
    libfreetype6 \
    ttf-mscorefonts-installer \
    && rm -rf /var/lib/apt/lists/*

RUN fc-cache -fv

WORKDIR /app

COPY publish/ .

RUN mkdir -p /output

ENTRYPOINT ["dotnet", "FRJ.Tools.SimpleWorkSheet.DockerTest.dll"]
