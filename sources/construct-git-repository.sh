#!/bin/bash

function _add {
  last_update="$1 $2"
  last_filename="$3"
  git add "$last_filename"
}
function _commit {
  local default_message="add program \"${last_filename%.*}\""
  local message=${1:-$default_message}
  git commit -m "$message" --date="$last_update"
}
function _commit_update {
  _commit "update $last_filename"
}
function _commit_add {
  _commit "add $last_filename"
}
function _merge {
  local branch=$1
  git merge "$branch"
  git branch -d "$branch"
}

function phase1 {
  _add 2001-06-17 15:04:28 信号.vbp
  _add 2001-06-17 15:04:28 信号.frm
  _add 2001-06-17 15:05:36 信号.vbw
  _commit

  _add 2001-06-23 22:17:20 Group1.vbg
  _commit

  _add 2001-07-28 13:17:24 牧場.frx
  _add 2001-07-28 13:17:26 牧場.frm
  _add 2001-07-28 13:22:20 牧場.vbp
  _commit
  _add 2001-08-04 14:37:28 牧場.vbw
  _commit_update

  _add 2001-08-04 15:32:06 色1.vbp
  _add 2001-08-04 15:32:06 KSColor.ctl
  _add 2001-08-04 15:32:14 色1.lib
  _add 2001-08-04 15:32:14 色1.exp
  _add 2001-08-04 15:32:16 色1.ocx
  _add 2001-08-04 15:32:42 色1.vbw
  _commit

  #_add 2001-08-07 22:08:52 MSCREATE.DIR
  _add 2001-08-15 20:42:32 電車.frx
  _add 2001-08-15 20:42:32 電車.FRM
  _add 2001-08-15 20:42:38 電車.vbg
  _commit

  _add 2001-09-24 18:46:54 飛行機.frx
  _add 2001-09-24 18:46:56 飛行機.vbp
  _add 2001-09-24 18:46:56 飛行機.frm
  _commit

  _add 2001-10-06 19:39:38 問題.frm
  _add 2001-10-06 19:39:40 問題だしプログラム.vbp
  _add 2001-10-06 20:31:10 問題だしプログラム.vbw
  _commit

  _add 2001-11-24 14:10:14 バウンド.frx
  _add 2001-11-24 14:10:14 バウンド.frm
  _add 2001-11-24 14:10:22 バウンド.vbp
  _commit

  _add 2001-12-08 16:31:52 ○×用.ctl
  _add 2001-12-08 16:53:30 ○×.vbw
  _add 2001-12-08 16:53:32 ○×Splash.frx
  _add 2001-12-08 16:53:32 ○×Splash.frm
  _add 2001-12-08 16:53:32 ○×.vbp
  _add 2001-12-08 16:53:32 ○×.frm
  _commit

  _add 2002-01-14 17:20:22 バウンド.vbw
  _commit_update

  _add 2002-01-14 17:35:08 卓球.frm
  _add 2002-01-14 17:35:18 卓球.vbp
  _add 2002-01-14 17:35:22 卓球.vbw
  _commit

  _add 2002-01-27 20:03:58 ％メーター.vbp
  _add 2002-01-27 20:04:54 ％メーター.ctl
  _commit

  _add 2002-01-27 20:05:18 試し１.frm
  _add 2002-01-27 20:05:32 試し１.vbp
  _add 2002-01-27 20:05:50 ％表示.vbg
  _add 2002-01-27 20:05:50 ％メーター.vbw
  _add 2002-01-27 20:05:52 試し１.vbw
  _commit

  _add 2002-02-02 17:30:30 π.vbp
  _add 2002-02-02 17:30:30 π.frm
  _add 2002-02-02 17:35:32 π.vbw
  _commit

  _add 2002-02-03 20:53:46 rpg1.frx
  _add 2002-02-03 20:53:48 rpg1.vbp
  _add 2002-02-03 20:53:48 rpg1.frm
  _add 2002-02-09 16:58:12 rpg1.vbw
  _commit

  _add 2002-02-24 14:34:24 shortcuts.frx
  _add 2002-02-24 14:34:26 shortcuts.vbw
  _add 2002-02-24 14:34:26 shortcuts.vbp
  _add 2002-02-24 14:34:26 shortcuts.frm
  _commit

  _add 2002-03-17 16:58:04 回転体aaa.vbw
  _add 2002-03-17 16:58:04 回転体aaa.vbp
  _add 2002-03-21 17:51:08 飛行機.vbw
  _commit

  _add 2002-03-30 15:11:10 ボールうち.exe
  _add 2002-03-30 15:15:18 ボールうち.vbw
  _add 2002-03-30 15:15:18 ボールうち.frx
  _add 2002-03-30 15:15:20 ボールうち.vbp
  _add 2002-03-30 15:15:20 ボールうち.frm
  _commit

  _add 2002-03-30 15:43:02 回転体.frx
  _add 2002-03-30 15:43:02 回転体.frm
  _commit

  _add 2002-04-13 21:45:56 カーレース.frx
  _add 2002-04-13 21:45:58 カーレース.vbp
  _add 2002-04-13 21:45:58 カーレース.frm
  _commit

  _add 2002-04-14 20:21:32 graph.vbw
  _add 2002-04-14 20:21:34 guraf.frm
  _add 2002-04-14 20:21:34 graph.vbp
  _commit

  _add 2002-04-27 14:55:54 陣１.frm
  _commit

  _add 2002-05-06 13:22:36 名前入力.frm
  _add 2002-05-06 13:24:04 陣３.frm
  _add 2002-05-06 14:02:50 陣取り.vbw
  _add 2002-05-06 14:02:52 陣取り.vbp
  _add 2002-05-06 14:02:52 陣２.frm
  _commit

  _add 2002-06-08 15:56:04 重力2.frm
  _add 2002-06-08 15:56:40 重力.frm
  _add 2002-06-08 15:57:00 重力.vbw
  _add 2002-06-08 15:57:02 重力.vbp
  _commit

  _add 2002-06-09 15:58:58 英単語保存.frm
  _add 2002-06-16 15:21:06 英単語.frm
  _add 2002-06-16 15:40:40 英単語開く.frm
  _add 2002-06-16 15:46:12 英単語出題.frx
  _add 2002-06-16 15:46:14 英単語出題.frm
  _add 2002-06-16 15:50:40 英単語出題.vbp
  _commit

  _add 2002-06-23 14:03:40 英単語登録.frm
  _add 2002-06-23 14:03:40 英単語主.frx
  _add 2002-06-23 14:03:40 英単語主.frm
  _add 2002-06-23 14:12:38 英単語.vbg
  _add 2002-06-23 14:13:02 英単語.vbw
  _add 2002-06-23 14:13:04 英単語出題.vbw
  _add 2002-06-23 14:13:04 英単語.vbp
  _commit

  _add 2002-07-06 16:43:44 π2.frx
  _add 2002-07-06 16:43:46 π2.vbw
  _add 2002-07-06 16:43:46 π2.vbp
  _add 2002-07-06 16:43:46 π2.frm
  _commit

  _add 2002-08-31 15:03:36 3d入力.frm
  _add 2002-08-31 15:44:18 3d表示.frx
  _add 2002-08-31 15:44:20 3d表示.frm
  _add 2002-08-31 15:48:02 三点透視.vbp
  _add 2002-08-31 15:48:06 三点透視.vbw
  _commit

  _add 2002-10-13 20:30:14 blind.frm
  _add 2002-10-13 20:30:24 blind.vbw
  _add 2002-10-13 20:30:24 blind.vbp
  _commit

  _add 2002-10-20 21:28:40 時計２.frx
  _add 2002-10-20 21:28:42 時計２.frm
  _add 2002-10-20 21:28:42 時計.frm
  _add 2002-10-20 21:30:38 時計.vbw
  _add 2002-10-20 21:30:38 時計.vbp
  _commit

  _add 2002-11-17 13:55:30 カーレース.vbw
  _commit

  _add 2002-12-21 15:26:48 アンノーン.vbw
  _add 2002-12-21 15:26:48 unknown.frm
  _add 2002-12-21 15:26:50 アンノーン.vbp
  _commit

  _add 2002-12-21 16:54:38 歴史年号.vbw
  _add 2002-12-21 16:54:38 歴史年号.vbp
  _add 2002-12-21 16:54:38 歴史年号.frm
  _commit

  _add 2003-03-20 20:11:14 表示unicodes.vbw
  _add 2003-03-20 20:11:14 表示unicodes.vbp
  _add 2003-03-20 20:11:14 表示unicodes.frm
  _commit

  _add 2003-03-26 22:19:52 ボタン.frm
  _commit

  _add 2003-03-27 09:52:04 ぱらぱら.vbw
  _add 2003-03-27 09:52:04 ぱらぱら.vbp
  _add 2003-03-27 09:52:04 ぱらぱら.frm
  _commit

  _add 2003-03-30 12:08:16 緊急コピー.vbp
  _add 2003-03-30 12:08:16 緊急コピー.frm
  _add 2003-03-30 23:14:56 緊急コピー.vbw
  _commit

  _add 2003-03-31 17:14:24 ekanji用.vbp
  _add 2003-03-31 17:14:24 ekanji用.frm
  _add 2003-03-31 17:14:32 ekanji用.vbw
  _commit

  _add 2003-04-02 21:32:10 "功filer s.frm"
  _commit_update

  _add 2003-04-05 13:10:16 絵1.frm
  _add 2003-04-06 18:26:04 絵道具.frm
  _add 2003-04-06 21:48:44 絵設定2.frm
  _add 2003-04-07 17:39:58 絵設定.frm
  _commit

  _add 2003-04-20 11:43:44 功filer.frx
  _add 2003-04-20 11:43:46 功filer.frm
  _add 2003-04-27 17:33:18 功filer.vbp
  _commit

  _add 2003-04-27 18:01:58 絵1.vbp
  _commit_update

  _add 2003-06-14 17:42:56 実験Form2.frm
  _add 2003-06-14 17:43:02 Project1.vbw
  _add 2003-06-14 17:43:02 Project1.vbp
  _commit

  _add 2003-06-21 17:25:40 回転体.vbp
  _commit_update

  _add 2003-06-21 17:28:22 絵1.vbw
  _commit_update

  _add 2003-06-21 17:28:36 功filer.vbw
  _add 2003-06-21 17:34:18 animatioon.vbw
  _commit_update

  _add 2003-06-21 19:08:22 回転体.vbw
  _commit_update

  _add 2003-06-21 19:11:40 功一計算.vbg
  _commit_update

  _add 2003-08-08 14:11:12 計算ctl.vbw
  _add 2003-08-08 14:11:12 計算2.vbw
  _add 2003-08-08 14:11:12 計算１.vbw
  _commit

  _add 2004-02-28 18:51:15 "vb7関数a to c.txt"
  _commit_add

  _add 2004-03-21 13:45:32 C#classやstruct説明.txt
  _commit_add

  _add 2005-08-28 13:00:59 desktop.ini
  _commit_add
}

function phase2 {
  git checkout -b lifegames 7d98382
  _add 2002-03-21 17:25:16 "lifegames/life Game.frm"
  _commit 'add program life1'
  _add 2002-06-08 14:48:10 lifegames/life2.frx
  _commit 'add program life2'
  _add 2003-04-12 19:55:30 lifegames/life3.frm
  _commit 'add program life3'
  _add 2003-07-13 13:06:54 lifegames/life4.vbp
  _add 2003-07-13 13:06:54 lifegames/life3.vbp
  _add 2003-07-13 13:06:54 lifegames/life2.vbp
  _add 2003-07-13 13:06:54 "lifegames/life r3.vbp"
  _add 2003-07-13 13:06:54 "lifegames/life Game.vbp"
  _add 2003-07-13 14:17:04 lifegames/SpinText.ctl
  _add 2003-07-13 14:29:52 lifegames/LifeGames.vbg
  _add 2003-07-13 14:29:52 lifegames/life4.frx
  _add 2003-07-13 14:29:52 lifegames/life4.frm
  _commit 'add program life4'
  _add 2003-08-02 15:53:24 lifegames/life4.vbw
  _add 2003-08-02 15:53:24 lifegames/life3.vbw
  _add 2003-08-02 15:53:24 lifegames/life2.vbw
  _add 2003-08-02 15:53:24 "lifegames/life r3.vbw"
  _add 2003-08-02 15:53:24 "lifegames/life Game.vbw"
  _commit 'update lifegames'
  _add 2004-03-13 11:33:40 lifegames/life2.frm
  _commit_update
}

function phase3 {
  git checkout -b Tetris ef03faae
  _add 2003-07-06 14:31:56 Tetris/tetris1.frm
  _commit_add
  _add 2003-07-06 15:36:12 Tetris/tetris.ico
  _commit_add
  _add 2003-07-06 16:28:52 Tetris/tetris2.frx
  _add 2003-07-06 16:28:52 Tetris/tetris2.frm
  _add 2003-07-06 16:28:52 Tetris/tetris1.vbw
  _add 2003-07-06 16:28:52 Tetris/tetris1.vbp
  _commit
}

function phase4 {
  _add 2003-06-22 13:53:24 充/test_1.vbw
  _add 2003-06-22 13:53:24 充/test_1.vbp
  _add 2003-06-22 13:53:24 充/test_1.frm
  _commit 'add project "test_1"'
}

function phase5 {
  git checkout -b keisan 0a5eab405
  _add 2003-06-14 16:40:46 計算/kmath.lib
  _add 2003-06-14 16:40:46 計算/kmath.exp
  _add 2003-06-14 16:40:48 計算/kmath.ocx
  _commit_add
  _add 2003-06-21 18:00:28 計算/描画関数.ctx
  _add 2003-06-21 18:00:30 計算/描画関数.ctl
  _commit_add
  _add 2003-06-21 18:56:02 計算/色関数ctl.ctx
  _add 2003-06-21 18:56:04 計算/色関数ctl.ctl
  _commit_add
  _add 2003-06-21 19:08:12 計算/計算ctl.ctx
  _commit_add
  _add 2003-08-08 14:01:12 計算/kmath.oca
  _add 2003-08-08 14:34:42 計算/計算ctl.vbp
  _commit_update
  _add 2003-08-09 18:26:58 計算/計算2.vbp
  _add 2003-08-09 18:37:38 計算/計算１.frx
  _add 2003-08-09 18:37:38 計算/計算１.frm
  _commit_add
  _add 2003-08-10 10:41:50 計算/logs.txt
  _add 2003-08-10 13:11:52 計算/計算2.frm
  _add 2003-08-10 13:11:52 計算/計算１.vbp
  _commit_update
  _add 2003-11-01 21:39:48 計算/計算2.ctx
  _add 2003-11-01 22:14:48 計算/units.ico
  _commit_add
  _add 2003-11-02 00:11:34 計算/単位の変換.frx
  _commit_add
  _add 2003-11-09 18:51:46 計算/三角形.frx
  _add 2003-11-09 18:51:46 計算/三角形.frm
  _commit_add
  _add 2003-11-09 18:52:38 計算/計算.vbg
  _add 2003-11-09 18:52:40 計算/計算ctl.vbw
  _add 2003-11-09 18:52:40 計算/計算2.vbw
  _add 2003-11-09 18:52:40 計算/計算１.vbw
  _commit_update
  _add 2004-01-24 15:07:29 計算/計算2.ctl
  _commit_update
  _add 2004-02-07 17:59:52 計算/計算ctl.ctl
  _add 2004-02-22 12:59:16 計算/単位の変換.frm
  _commit_add
}

function phase6 {
  git checkout -b ongaku 0a5eab405
  _add 2003-05-11 18:56:58 音楽/音楽1.vbg
  _commit
  _add 2003-07-20 16:01:54 音楽/onngaku1.log
  _add 2003-08-01 18:50:06 音楽/音楽1設定.frx
  _add 2003-08-01 18:50:06 音楽/音楽1設定.frm
  _commit_add
  _add 2003-08-09 12:56:02 音楽/音楽1.frm
  _add 2003-08-09 12:58:24 音楽/onngaku1.frx
  _add 2003-08-09 12:58:24 音楽/onngaku1.frm
  _add 2003-08-09 13:03:24 音楽/音楽１.vbp
  _commit_update
  _add 2003-11-30 15:59:50 音楽/音楽１.vbw
  _commit_update
}

function phase7 {
  git checkout -b molecular ff6fc8555
  mv 分子模型.NET/original 分子模型
  _add 2002-09-21 19:06:58 分子模型/分子模型.frm
  _commit
  _add 2002-10-06 10:26:34 分子模型/見る角度.frm
  _commit_add
  _add 2002-10-06 11:10:14 分子模型/分子模型frm.frm
  _add 2002-10-06 11:10:14 分子模型/分子模型.vbw
  _add 2002-10-06 11:10:14 分子模型/分子模型.vbp
  _commit_update

  mv 分子模型 分子模型.NET/original
  _add 2004-06-12 22:02:09 分子模型.NET/AssemblyInfo.vb
  _add 2004-06-12 22:02:10 分子模型.NET/AxSpinButtonArray.vb
  _add 2004-06-12 22:02:10 分子模型.NET/分子模型frm.resX
  _add 2004-06-12 22:02:10 分子模型.NET/分子模型.resX
  _add 2004-06-12 22:02:10 分子模型.NET/見る角度.resX
  _add 2004-06-12 22:02:11 分子模型.NET/分子模型.vb
  _add 2004-06-12 22:02:29 分子模型.NET/AxSpinButtonArray.dll
  _add 2004-06-12 22:02:31 分子模型.NET/分子模型.log
  _add 2004-06-12 22:17:45 分子模型.NET/見る角度.vb
  _add 2004-06-12 22:17:48 分子模型.NET/分子模型frm.vb
  _add 2004-06-12 22:17:48 分子模型.NET/分子模型.vbproj
  _add 2004-06-12 22:18:04 分子模型.NET/分子模型.sln
  _commit
}

function phase8 {
  git checkout -b zangai 8b0882f9c
  _add 2001-09-15 21:14:16 ？電卓の残骸/Form1.frm
  _commit 'add 電卓の残骸'
  _add 2003-06-28 19:59:40 ？電卓の残骸/Project1.vbp
  _add 2003-06-28 20:01:38 ？電卓の残骸/Form1.log
  _add 2003-06-28 20:05:12 ？電卓の残骸/Project1.vbw
  _commit 'update 電卓の残骸'
}

function phase9 {
  git checkout -b machi 93e7a9d3f
  _add 2001-08-10 13:55:28 町/町.vbw
  _add 2001-08-10 13:55:28 町/町.frx
  _add 2001-08-10 13:55:30 町/町.vbp
  _add 2001-08-10 13:55:30 町/町.frm
  _commit
  _add 2005-01-15 22:54:59 町/a.bin
  _add 2005-01-16 02:09:55 町/x.bmp
  _add 2005-01-16 02:18:42 町/x.ico
  _add 2005-01-16 02:24:15 町/frx.txt
  _commit 'add resources'
  _add 2005-01-17 01:09:13 町/町.txt
  _add 2005-01-17 01:09:14 町/test.txt
  _commit 'add files'
}


function phase10 {
  git checkout -b anntena ff6fc85
  _add 2002-09-16 18:10:30 VB_ANNTENA/Form1.frm
  _add 2002-09-16 18:22:40 VB_ANNTENA/Project1.vbw
  _add 2002-09-16 18:22:42 VB_ANNTENA/Project1.vbp
  _add 2002-09-16 18:22:42 VB_ANNTENA/Form2.frm
  _commit 'add project "VB_ANNTENA"'

  git checkout -b kensaku 83d7cfe8
  _add 2004-01-03 19:59:48 検索/検索.frx
  _add 2004-01-03 20:50:56 検索/検索.vbp
  _add 2004-01-03 20:50:58 検索/検索.vbw
  _add 2004-06-13 00:27:26 検索/検索.frm
  _commit

  git checkout -b binary 7d89af42
  _add 2005-01-01 18:04:46 binary/binary_main.frx
  _add 2005-01-01 18:04:46 binary/binary_main.frm
  _add 2005-01-01 18:05:34 binary/binary.bas
  _add 2005-01-01 18:07:46 binary/binary_read.bas
  _add 2005-01-01 18:07:54 binary/binary.vbp
  _add 2005-01-01 18:07:56 binary/binary.vbw
  _commit
}

function phase11 {
  git checkout -b tameshi 0a5eab40
  _add 2003-03-26 10:47:52 試しプログラム/desktop.ini
  _add 2003-05-04 16:31:00 試しプログラム/form.vbs
  _commit
}

function phase12 {
  git checkout ongaku
  (
    cd ..
    _add 1998-06-24 00:00:00 音楽/TABCTL32.OCX
    _add 2000-05-22 16:58:00 音楽/MCI32.OCX
    _add 2002-09-07 11:26:58 音楽/a7.wav
    _add 2002-09-07 11:27:36 音楽/a8.wav
    _add 2002-09-07 11:28:20 音楽/a9.wav
    _add 2002-09-07 11:28:46 音楽/a10.wav
    _add 2002-09-07 11:29:12 音楽/a11.wav
    _add 2002-09-07 11:29:32 音楽/a12.wav
    _add 2002-09-07 11:30:36 音楽/a13.wav
    _add 2002-09-07 11:31:04 音楽/a14.wav
    _add 2002-09-07 11:31:38 音楽/a15.wav
    _add 2002-09-07 11:32:10 音楽/a16.wav
    _add 2002-09-07 11:33:06 音楽/a17.wav
    _add 2002-09-07 11:33:44 音楽/a18.wav
    _add 2002-09-07 11:34:44 音楽/a19.wav
    _add 2002-09-08 10:58:38 音楽/a20.wav
    _add 2002-09-08 11:00:20 音楽/a21.wav
    _add 2002-09-08 11:00:56 音楽/a22.wav
    _add 2002-09-08 11:02:32 音楽/a23.wav
    _add 2002-09-08 11:02:58 音楽/a24.wav
    _add 2002-09-08 11:05:46 音楽/a25.wav
    _add 2002-09-08 11:06:42 音楽/a26.wav
    _add 2002-09-08 11:07:08 音楽/a27.wav
    _add 2002-09-08 11:07:28 音楽/a28.wav
    _add 2002-09-08 11:07:54 音楽/a29.wav
    _add 2002-09-08 11:08:16 音楽/a30.wav
    _add 2002-09-08 11:09:52 音楽/a31.wav
    _add 2002-09-08 12:39:10 音楽/a0.wav
    _add 2002-09-08 12:39:34 音楽/a1.wav
    _add 2002-09-08 12:39:52 音楽/a2.wav
    _add 2002-09-08 12:40:16 音楽/a3.wav
    _add 2002-09-08 12:40:54 音楽/a4.wav
    _add 2002-09-08 12:41:14 音楽/a5.wav
    _add 2002-09-08 12:41:32 音楽/a6.wav
    _add 2002-09-07 11:26:58 音楽/wavdat/a7.wav
    _add 2002-09-07 11:27:36 音楽/wavdat/a8.wav
    _add 2002-09-07 11:28:20 音楽/wavdat/a9.wav
    _add 2002-09-07 11:28:46 音楽/wavdat/a10.wav
    _add 2002-09-07 11:29:12 音楽/wavdat/a11.wav
    _add 2002-09-07 11:29:32 音楽/wavdat/a12.wav
    _add 2002-09-07 11:30:36 音楽/wavdat/a13.wav
    _add 2002-09-07 11:31:04 音楽/wavdat/a14.wav
    _add 2002-09-07 11:31:38 音楽/wavdat/a15.wav
    _add 2002-09-07 11:32:10 音楽/wavdat/a16.wav
    _add 2002-09-07 11:33:06 音楽/wavdat/a17.wav
    _add 2002-09-07 11:33:44 音楽/wavdat/a18.wav
    _add 2002-09-07 11:34:44 音楽/wavdat/a19.wav
    _add 2002-09-08 10:58:38 音楽/wavdat/a20.wav
    _add 2002-09-08 11:00:20 音楽/wavdat/a21.wav
    _add 2002-09-08 11:00:56 音楽/wavdat/a22.wav
    _add 2002-09-08 11:02:32 音楽/wavdat/a23.wav
    _add 2002-09-08 11:02:58 音楽/wavdat/a24.wav
    _add 2002-09-08 11:05:46 音楽/wavdat/a25.wav
    _add 2002-09-08 11:06:42 音楽/wavdat/a26.wav
    _add 2002-09-08 11:07:08 音楽/wavdat/a27.wav
    _add 2002-09-08 11:07:28 音楽/wavdat/a28.wav
    _add 2002-09-08 11:07:54 音楽/wavdat/a29.wav
    _add 2002-09-08 11:08:16 音楽/wavdat/a30.wav
    _add 2002-09-08 11:09:52 音楽/wavdat/a31.wav
    _add 2002-09-08 12:39:10 音楽/wavdat/a0.wav
    _add 2002-09-08 12:39:34 音楽/wavdat/a1.wav
    _add 2002-09-08 12:39:52 音楽/wavdat/a2.wav
    _add 2002-09-08 12:40:16 音楽/wavdat/a3.wav
    _add 2002-09-08 12:40:54 音楽/wavdat/a4.wav
    _add 2002-09-08 12:41:14 音楽/wavdat/a5.wav
    _add 2002-09-08 12:41:32 音楽/wavdat/a6.wav
    _commit '音楽: add sound data'
    _add 2003-07-22 17:28:46 音楽/16.cur
    _add 2003-07-22 17:29:00 音楽/4.cur
    _add 2003-07-22 17:29:08 音楽/8.cur
    _add 2003-07-22 17:29:16 音楽/2.cur
    _add 2003-07-22 17:29:30 音楽/6.cur
    _add 2003-07-22 17:29:38 音楽/12.cur
    _add 2003-07-22 17:29:42 音楽/3.cur
    _add 2003-07-22 17:35:20 音楽/1.cur
    _commit '音楽: add cursors'
    _add 2003-07-26 13:57:10 音楽/0.cur
    _commit 'add a cursor for 音楽'
    _add 2003-07-27 18:38:24 音楽/音符.bmp
    _add 2003-07-27 18:47:12 音楽/速さ.bmp
    _commit '音楽: add a images'
    _add 2003-08-03 19:45:32 音楽/play.ico
    _add 2003-08-03 19:47:00 音楽/play2.ico
    _add 2003-08-03 19:48:08 音楽/end.ico
    _add 2003-08-03 19:49:22 音楽/stop.ico
    _add 2003-08-09 11:46:24 音楽/98.cur
    _add 2003-08-09 11:47:36 音楽/99.cur
    _commit '音楽: add icons/cursors'
  )
}

function phase13 {
  git checkout -b data d5a64ed3
  (
    cd ..
    _add 2003-07-26 20:35:38 ボールうち/Desktop.ini
    _add 2003-07-21 14:03:54 音楽/Desktop.ini
    _add 2002-01-27 19:34:50 ボールうち/"read me.txt"
    _commit 'add directories'
    _add 2002-06-16 14:19:36 英単語/書.txt
    _add 2002-06-16 15:58:02 英単語/何か.txt
    _commit_add
    _add 2002-06-23 14:14:52 英単語/総合.txt
    _commit_add
    _add 2002-07-14 18:03:56 功一電卓/質量.単位１
    _commit_add
    _add 2002-07-21 10:19:00 unknown/mituru.txt
    _add 2002-07-22 14:02:22 unknown/無題.txt
    _commit_add
    _add 2002-09-07 16:48:18 音楽/小狐コンコン.kon
    _add 2002-09-07 16:54:50 音楽/てふてふ.kon
    _add 2002-09-07 17:14:04 音楽/きらきら星.kon
    _add 2002-09-07 17:14:42 音楽/結んで開いて.kon
    _add 2002-09-07 17:24:26 音楽/雪やこんこ.kon
    _add 2002-09-07 17:34:08 音楽/春の小川.kon
    _commit_add
    _add 2002-09-15 15:50:18 音楽/風になれ.kon
    _commit_add
    _add 2002-09-16 15:23:56 音楽/海の不思議.kon
    _commit_add
    _add 2002-11-23 15:01:52 音楽/心の中にきらめいて.kon
    _commit_add
    _add 2002-12-14 17:16:40 英単語/試し.txt
    _commit_add
    _add 2003-03-17 15:54:10 功一電卓/面積.単位１
    _add 2003-03-17 16:02:28 功一電卓/体積.単位１
    _add 2003-03-17 17:40:02 功一電卓/長さ.単位１
    _commit_add
    _add 2003-03-19 16:38:26 功一電卓/meter.単位２a
    _commit_add
    _add 2003-03-20 11:40:42 功一電卓/量2.単位２a
    _add 2003-03-20 20:47:56 功一電卓/meter3.単位２a
    _add 2003-03-20 21:11:46 功一電卓/度・長さ.単位２
    _commit_add
    _add 2003-03-21 08:18:04 功一電卓/面積・平方尺.単位２a
    _add 2003-03-21 08:18:04 功一電卓/面積・方歩.単位２a
    _add 2003-03-21 08:18:04 功一電卓/面積・畝.単位２a
    _commit_add
    _add 2003-03-24 18:17:40 功一電卓/衡・質量.単位２
    _commit_add
    _add 2003-03-25 17:06:48 功一電卓/meter2.単位２a
    _add 2003-03-25 17:10:38 功一電卓/方歩.単位２a
    _add 2003-03-25 17:49:38 功一電卓/畝.単位２a
    _add 2003-03-25 18:22:48 功一電卓/.計算1
    _add 2003-03-25 18:23:26 功一電卓/aaa.計算1
    _commit_add
    _add 2003-04-12 21:24:14 功一電卓/充.計算1
    _commit_add
    _add 2003-04-12 21:39:06 英単語/きり.txt
    _commit_add
    _add 2003-04-26 17:20:36 音楽/千種台中学校校歌.kon
    _add 2003-04-26 18:07:36 音楽/旭丘高等学校校歌.kon
    _commit_add
    _add 2003-04-27 18:10:22 音楽/南毛利小学校校歌.kon
    _add 2003-04-27 18:16:32 音楽/茶摘み.kon
    _add 2003-04-27 18:21:36 音楽/赤とんぼ.kon
    _add 2003-04-27 18:26:34 音楽/エーデルワイス.kon
    _commit_add
    _add 2003-05-04 14:42:12 音楽/月の光に.kon
    _add 2003-05-04 15:00:22 音楽/おぼろ月夜.kon
    _commit_add
    _add 2003-05-10 15:01:04 "音楽/top of the world.kon"
    _add 2003-05-10 15:05:34 音楽/充original.kon
    _commit_add
    _add 2003-05-11 17:41:16 音楽/カエルの歌.kon
    _commit_add
    _add 2003-07-20 16:17:00 音楽/南中卒業式　メドレー2.kon
    _add 2003-07-20 16:17:34 音楽/南中卒業式　メドレー1.kon
    _add 2003-07-20 16:19:08 音楽/南中卒業式　メドレー3.kon
    _commit_add
    _add 2003-07-21 14:03:54 音楽/Desktop.ini
    _add 2003-07-21 14:04:32 功一電卓/Desktop.ini
    _commit_add
    _add 2003-07-21 18:53:28 音楽/少年時代.kon
    _add 2003-07-21 19:22:54 音楽/世界って広いわ.kon
    _commit_add
    _add 2003-07-25 14:37:12 功一電卓/面積.単位２
    _add 2003-07-25 14:43:04 功一電卓/衡.単位２a
    _add 2003-07-25 14:43:32 功一電卓/度.単位２a
    _add 2003-07-25 14:45:30 功一電卓/量1.単位２a
    _commit_add
    _add 2003-08-01 18:52:02 音楽/blank.kon
    _add 2003-08-01 18:52:50 "音楽/春  A.Vivaldi.kon"
    _commit_add
    _add 2003-08-02 17:11:50 音楽/元気になれそう.kon
    _commit_add
    _add 2003-08-09 15:31:42 音楽/心の瞳.kon
    _add 2003-08-09 16:30:56 功一電卓/量・体積.単位２
    _commit_add
    _add 2003-10-25 21:37:22 功一電卓/？.zip
    _add 2003-10-25 21:40:44 功一電卓/旧.zip
    _commit_add
    _add 2004-05-02 15:25:00 "音楽/散歩 中川李枝子.kon"
    _commit_add
    _add 2004-05-30 18:04:01 功一電卓/九九表.txt
    _commit_add
    _add 2005-07-24 17:18:01 音楽/南毛利中学校校歌.kon
    _commit_add
    _add 2008-03-08 10:47:16 功一電卓/平方根表.txt
    _commit_add
  )
}

#git merge binary kensaku Tetris keisan tameshi ongaku molecular anntena lifegames zangai machi
function phase14 {
  _merge binary
  _merge kensaku
  _merge Tetris
  _merge keisan
  _merge tameshi
  _merge ongaku
  _merge molecular
  _merge anntena
  _merge lifegames
  _merge zangai
  _merge machi
  _merge data
}
