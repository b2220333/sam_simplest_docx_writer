echo '開始commit新版本：'$1
git add *
git commit -m $1
git push

