go build -o ./build/go-ej main.go
CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -o ./build/go-ej.exe main.go