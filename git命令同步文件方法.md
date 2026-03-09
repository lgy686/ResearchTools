# 使用git命令来同步vscode中的文件
## 初次上传
```
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/lgy686/仓库名称需要更换
git push -u origin main
```
## 某个文件发生改动后需要更新
```
git status
git add 发生改动的文件名称
git commit -m "update searchletpub"
git pull --rebase origin main
git push origin main
```
