.PHONY: build run test docker-build docker-up docker-down lint clean

APP_NAME := docx-api
PORT     := 8080

## build: Compile the binary
build:
	go build -o bin/$(APP_NAME) ./cmd/server

## run: Run the server locally
run: build
	PORT=$(PORT) ./bin/$(APP_NAME)

## test: Run all tests
test:
	go test -v -race -count=1 ./...

## test-cover: Run tests with coverage
test-cover:
	go test -v -race -coverprofile=coverage.out ./...
	go tool cover -func=coverage.out

## lint: Run linter
lint:
	golangci-lint run ./...

## docker-build: Build the Docker image
docker-build:
	docker compose build

## docker-up: Start with Docker Compose
docker-up:
	docker compose up -d

## docker-down: Stop Docker Compose
docker-down:
	docker compose down

## docker-logs: Show container logs
docker-logs:
	docker compose logs -f docx-api

## clean: Remove build artifacts
clean:
	rm -rf bin/ coverage.out tmp/

## help: Show this help
help:
	@grep -E '^## ' Makefile | sed 's/## //' | column -t -s ':'
