# pyopentd

（開発中）

Thermal DesktopのAPIであるOpenTDをpythonで触るためのライブラリです。
公式のチュートリアルではC#がメインで、pythonでの例が非常に少なく苦労したので、ライブラリを作成しました。

使用シチュエーションとしては、元々あるTDモデルから、いくつかの変数（軌道やシンボル等）を変更して実行し、結果を検証する、という使い方を想定しています。ゼロからモデルを作ることは現状想定していません。

## ライブラリ

（※後日丁寧に書きます）

- python = 3.8 (3.9以降はpythonnetが対応していないらしい。3.8以下であれば大丈夫？)
- pythonnet
- clr
- pandas
- numpy

\* OpenTD : Version 6.2

## はじめに（インストール・環境構築）

（※後日丁寧に書きます）

1. ダウンロード

    ``` PowerShell
    git clone https://github.com/nishitomo206/pyopentd.git
    ```

2. ./pyopentd/src/pyopentd/ディレクトリにある、\_\_init\_\_.pyを除く全pythonファイルの"/OpenTDv62/v4.0_6.2.0.7__65e6d95ed5c2e178/"の行を自分のものに変更する。
3. windows powershellを開いて元のディレクトリ（README.mdと同じ階層）に移動。

    ``` PowerShell
    cd pyopentd
    ```

4. pip installする。

    ``` PowerShell
    pip install .
    ```

## サンプルを実行

./sample/sample.ipynb がサンプルファイルになっているので、ファイルを開いて実行してみて下さい。
