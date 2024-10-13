FROM --platform=linux/arm64/v8 golang:1.22

WORKDIR /go/src

# 必要なツールと依存関係をインストール
RUN apt-get update && apt-get install -y git curl

# air をダウンロードしてインストール
RUN curl -sSfL https://raw.githubusercontent.com/cosmtrek/air/master/install.sh | sh -s -- -b /go/bin

# プロジェクトのルートディレクトリにある go.mod と go.sum をコピー
COPY go.mod go.sum ./

# 依存関係をインストール
RUN go mod download

# ソースコードと設定ファイルをコピー
COPY . .

# air.toml ファイルをコピー
COPY .air.toml /go/src/.air.toml

# デフォルトコマンドを設定
CMD ["air", "-c", "/go/src/.air.toml"]
