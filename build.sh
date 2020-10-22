go build -o excelToJson main.go
CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -o excelToJson.exe main.go