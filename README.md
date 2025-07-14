# The sample of using an `.xlsx` file as a database

## Clone and move the project directory

```
git clone git@github.com:ega1979/excel_db_sample.git 
cd excel_db_sample
```

## Install python 3.x

```
$ brew install pyenv
$ pyenv install 3.12.10
$ pyenv local 3.12.10 # 当該フォルダのみ3.12.10に変更
$ python --version # → Python 3.12.10 と出ればOK
```

### Note
pyenv が正しく機能するには、.bashrc または .zshrc に以下の設定が必要です（Homebrew経由なら自動で入るはずですが確認を）：

```
export PYENV_ROOT="$HOME/.pyenv"
export PATH="$PYENV_ROOT/bin:$PATH"
eval "$(pyenv init --path)"
eval "$(pyenv init -)"
```

設定後はシェルを再起動するか、source ~/.zshrc or source ~/.bashrcなどを実行

### .zsh or .bash?

```
$ echo $SHELL
/bin/zsh # zshなので~/.zshrc
/bin/bash # bashなので~/.bashrc（または ~/.bash_profile
```

## Create `ticket.xlsx` as database.

1. 同じフォルダ内にticket.xlsxを作成する

2. Sheet1に以下をコピペ

| チケット番号 | 名前       | 所属 | 出席 |
|--------------|------------|------|------|
| 10001        | 山田恵一   | A社  |      |
| 10002        | 鈴木早智子   | B社  |      |
| 10003        | 田中八郎   | C社  |      |
| 10004        | 広瀬まなつ | D社  |      |
| 10005        | 秋吉昭二   | E社  |      |
| 10006        | 頓所隼   | F社  |      |
| 10007        | 荒川真二   | G社  |      |

## Run Python script

```
$ python excel_test.py
```
