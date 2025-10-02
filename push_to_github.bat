@echo off
cd /d C:\Users\M\source\repos\MultiToolWin

echo ================================
echo 推送 MultiToolWin 到 GitHub...
echo ================================

:: 确保分支名字是 main
git branch -M main

:: 设置远程仓库地址（替换成你的 GitHub 仓库地址）
git remote set-url origin https://github.com/475416861/MultiToolWin.git

:: 提交所有修改
git add .
git commit -m "更新代码"

:: 推送到 GitHub
git push -u origin main

echo ================================
echo 推送完成，请检查 GitHub 页面
echo https://github.com/475416861/MultiToolWin
echo ================================
pause
