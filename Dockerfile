FROM golang:1.25-alpine

RUN apk add --no-cache git ca-certificates

RUN go install github.com/githubnemo/CompileDaemon@latest

WORKDIR /src

COPY go.mod go.sum ./
COPY docx-go/go.mod ./docx-go/

RUN go mod download

COPY . .

RUN mkdir -p /tmp/docx-uploads

EXPOSE 8080

COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

ENTRYPOINT ["/entrypoint.sh"]