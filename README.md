# bpac_python-sample

## セットアップ
```
# 任意のGit用ディレクトリへ移動
>cd path\to\dir

# GitHubからカレントディレクトリへclone
path\to\dir>git clone https://github.com/thinkAmi-sandbox/bpac_python-sample.git

# virtualenv環境の作成とactivate
# *Python3.5は、`c:\python35-32\`の下にインストール
path\to\dir>virtualenv -p c:\python35-32\python.exe env
path\to\dir>env\Scripts\activate

# requirements.txtよりインストール
(env)path\to\dir>pip install -r requirements.txt

# 実行
## pythonnetを使う場合
(env)path\to\dir>python pythonnet_ver.py
```

　  
## テスト環境

- Windows10
- b-PAC SDK 3.1.004 (x86版)
- Python 3.5.1
- pythonnet 2.1.0
- pywin32 220


　  
## 関係するブログ

- [b-PAC SDKをPython + pythonnetで操作してみた - メモ的な思考的な](http://thinkami.hatenablog.com/entry/2016/05/10/220912)
- [b-PAC SDKをPython + pywin32(win32com)で操作してみた - メモ的な思考的な](http://thinkami.hatenablog.com/entry/2016/05/11/061626)