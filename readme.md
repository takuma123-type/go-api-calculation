go projectをbuild

- 1. ルートディクトリで`cd docker`コマンドを打つ
- 2. `docker compose build`でbuildをする
- 3. `docker compose run go sh`コマンドを打ち、コンテナ内に入る。
- 4. `go run main.go`コマンドを打ち、実行させる。

main.goを編集したら、 `go run main.go`コマンドを打ち、実行させること。
また、docker descktopを起動させること。