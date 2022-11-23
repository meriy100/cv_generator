# CV generator

FireStore から API 経由でデータを取得し､ 職務経歴書を作成する GAS プロジェクト

## Configuration 

GCP Console から 認証情報 "OAuth 2.0 クライアント ID" を作成し, credential ファイル を client_secret.json に配 

```sh
$ clasp setting projectId ${PROJECT_ID}
$ clasp login --creds client_secret.json
$ clasp pull ${GAS PROJECT ID} # .clasp.json を編集する必要があるかも 
```

```
# ローカル実行
$ clasp run main
```

```
# CAS プロジェクトへの push 
$ clasp push
```
