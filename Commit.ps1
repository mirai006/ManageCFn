

$DateTime = (Get-Date -Format "yyyy/MM/dd HH:mm") + " (JST)"

echo "`ncscript vbac.wsf decombine`n"
cscript vbac.wsf decombine

echo "`ngit add`n"
git add .

echo "`ngit commit`n"
git commit -m "$args Commited at $DateTime"

echo "`ngit push master`n"
git push
