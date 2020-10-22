go build -o ./build/excelToJson main.go
CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -o ./build/excelToJson.exe main.go