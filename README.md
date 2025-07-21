# name
ComPort Sniffer

![screenShot](https://github.com/user-attachments/assets/94dd72dd-9f38-4146-bc8c-344b6f7b6adc)


## Overview
シリアル通信している機器の間にPCを割り込みさせ通信データをモニタするツールです。

PCをオンラインモニタとして機能させます。

## Requirement
Windows PowerShell

## Usage
２個のCOMポートが使用できるPCにて、それぞれの機器を接続します。

機器の通信設定を確認しツールを起動します。

●起動オプション パラメータ

    ComPortSniffer.exe COM1 38400 None 8 One COM2 38400 None 8 One -byte -log

    COM1：COM番号

    38400：ボーレートbps

    None：パリティ（Even、Mark、None、Odd、Space）

    8：データ長（8、7、6、5）

    One：ストップビット（None、One、OnePointFive、Two）

    -byte：バイナリ値も表示。省略可能。

    -log： ログファイルを出力。省略可能。

ツール起動用フォームを使うと簡単に通信設定ができます。

![screenShot2](https://github.com/user-attachments/assets/170d7854-6f62-4474-a53c-9635a3acc6d9)

コマンドプロンプト画面に通信データが表示されます。

## Features
ツールの動作としては、一方のCOMポートで受信したデータをもう一方のCOMポートへそのまま送信するだけです。

データはPC上で一旦バッファリングされ出力しているため、原理的には通信速度が低下します。

互いの機器が異なる通信設定であっても使用が可能であり、設定変更が難しい機器での動作テストができます。

また、通常のオンラインモニタでは難しいUSB接続（仮想COM）でも工夫をすれば対応できます。

注意：フロー制御はしていませんので大量のデータを通信するとデータ抜けが発生する場合があります。

## Licence
The source code is licensed MIT. The website content is licensed CC BY 4.0,see LICENSE.
