# mirai project

## Manage CloudFormation Tool

## Commit 
Execute commit.ps1 [Comments]

## Pull
Only /bin/*.xlsm
If you want to update /src/*.bas , Execute cscript vbac.wsf decombine.

https://tonari-it.com/vba-vbac-git/

#  直したいところ
External Valueを引っ張るとき、時分のExcelから作った値は採用しない
External　Balueを作るとき、以前のファイルは削除する

## ちょっと難易度が高い修正
無限ループに入るJsonがある
リソースにStringのみの登録があると、エラーとなる
