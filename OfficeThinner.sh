#!/bin/bash

# This script is only for Office 2016 for Mac.
# Tested on Office version 15.16 (151105)
# Tested on Office version 15.17 (151206)
# Tested on Office version 15.18 (160109)
# Use at your own risk.

# Some large-sized duplicate files exist in Word.app, Excel.app, PowerPoint.app, Outlook.app and OneNote.app in Office 2016 for Mac.
# It's wasting your precious HDD/SSD space as these files are just 5 duplicate copies in these apps.
# This script moves some of the duplicate files from Excel, PowerPoint, Outlook and OneNote into a backup directory, and then soft link neccesary paths back to the real files in Word.app alone.

# Author: goodbest
# Repo: https://github.com/goodbest/OfficeThinner


# Configs. You can change path here
basePATH="/Applications/"
backupPATH="$HOME/Desktop/OfficeThinnerBackup/"

WordPATH="Microsoft Word.app"
ExcelPATH="Microsoft Excel.app"
PowerPointPATH="Microsoft PowerPoint.app"
OutlookPATH="Microsoft Outlook.app"
OneNotePATH="Microsoft OneNote.app"

# Symbolic link the files in following 4 apps to Word.app.
# If an app's version is not same with Word, it will be excluded.
versionPATH="/Contents/Info.plist"
versionKey="CFBundleShortVersionString"
wordVersion=$(defaults read "$basePATH$WordPATH$versionPATH" $versionKey)

appPathArrayPending=( "$ExcelPATH" "$PowerPointPATH" "$OutlookPATH" "$OneNotePATH" )
appPathArray=()
for appPATH in "${appPathArrayPending[@]}";
do
	if [ -d "$basePATH$appPATH" ]; then
		appVersion=$(defaults read "$basePATH$appPATH$versionPATH" $versionKey)
		if [ $wordVersion == $appVersion ]; then
			appPathArray+=("$appPATH")
		else
			echo "WARNING: WILL NOT deal with ${appPATH/.app/}. It is version $appVersion, but Word is $wordVersion"
		fi
	else
		echo "WARNING: WILL NOT deal with ${appPATH/.app/}. It is NOT installed."
	fi
done
for appPATH in "${appPathArray[@]}";
do
     echo "WILL deal with ${appPATH/.app/}."
done

# If all apps are excluded, just exit the script.
if [ ${#appPathArray[@]} -eq 0 ]; then
    echo "No app will be dealed. Bye"
    exit
fi

# Add y/n choice.
read -n1 -r -p "Do you want to continue? y/n..." key < /dev/tty
if [ "$key" != 'y' ]; then
    if [ "$key" != 'Y' ]; then
        echo ""
        echo "Terminated. Bye"
        exit
    fi
fi

# Disk Usage Display
diskUsage(){
    for appPATH in "${@}";
    do
        du -sh "$basePATH$appPATH"
    done
    echo ""
}

echo ""
echo ""
echo "Before running this script, Office is taking:"
diskUsage "${appPathArray[@]}"

# ==============================
# Phase I: Deal with Fonts
# Comparison Result:  Word = Excel = Powerpoint, Outlook is subset of Word, OneNote is subset of Word
# ==============================
fontPATH="/Contents/Resources/"
if [ -d "$basePATH$WordPATH$fontPATH/DFonts" ]; then
    fontName="DFonts"
else
    fontName="Fonts"
fi
# echo "Thinning Fonts, it saves you ~1.4G space"
for appPATH in "${appPathArray[@]}";
do
    appName=${appPATH/.app/}
    mkdir -p "$backupPATH$appName$fontPATH"
    # echo "$backupPATH$appName$fontPATH"
    sudo mv "$basePATH$appPATH$fontPATH$fontName" "$backupPATH$appName$fontPATH"
    sudo ln -s "$basePATH$WordPATH$fontPATH$fontName" "$basePATH$appPATH$fontPATH$fontName"
    # echo "$basePATH$appPATH$fontPATH$fontName" "$backupPATH$appName$fontPATH"
    # echo "$basePATH$WordPATH$fontPATH$fontName" "$basePATH$appPATH$fontPATH$fontName"
done

# ==============================
# Phase II: Deal with Proofing Tools
# Comparison Result:  Word = Excel = Powerpoint = Outlook, OneNote is subset of Word
# ==============================
proofingPATH="/Contents/SharedSupport/"
proofingName="Proofing Tools"
# echo "Thinning Proofing Tools, it saves you ~1.5G space"
for appPATH in "${appPathArray[@]}";
do
    appName=${appPATH/.app/}
    mkdir -p "$backupPATH$appName$proofingPATH"
    # echo "$backupPATH$appName$proofingPATH"
    sudo mv "$basePATH$appPATH$proofingPATH$proofingName" "$backupPATH$appName$proofingPATH"
    sudo ln -s "$basePATH$WordPATH$proofingPATH$proofingName" "$basePATH$appPATH$proofingPATH$proofingName"
    # echo "$basePATH$appPATH$proofingPATH$proofingName" "$backupPATH$appName$proofingPATH"
    # echo "$basePATH$WordPATH$proofingPATH$proofingName" "$basePATH$appPATH$proofingPATH$proofingName"
done

# ==============================
# Phase III: Deal with MicrosoftOffice.framework
# Comparison Result:  Word = Excel = Powerpoint = Outlook = OneNote
# Maybe it's not a good idea to symbolic link this.
# ==============================
# frameworkPATH="/Contents/Frameworks/MicrosoftOffice.framework/Versions/A/"
# frameworkName="Resources"
# echo "Thinning MicrosoftOffice.framework, it saves you ~0.8G space"
# for appPATH in "${appPathArray[@]}";
# do
    # appName=${appPATH/.app/}
    # mkdir -p "$backupPATH$appName$frameworkPATH"
    # echo "$backupPATH$appName$frameworkPATH"
    # sudo mv "$basePATH$appPATH$frameworkPATH$frameworkName" "$backupPATH$appName$frameworkPATH"
    # sudo ln -s "$basePATH$WordPATH$frameworkPATH$frameworkName" "$basePATH$appPATH$frameworkPATH$frameworkName"
    # echo "$basePATH$appPATH$frameworkPATH$frameworkName" "$backupPATH$appName$frameworkPATH"
    # echo "$basePATH$WordPATH$frameworkPATH$frameworkName" "$basePATH$appPATH$frameworkPATH$frameworkName"
# done


echo "After running this script, Office is taking:"
diskUsage "${appPathArray[@]}"

echo "Office Thinning Complete."
echo "The duplicate files are backed up at $backupPATH"
echo "If everything is OK, you may delete these files. But the choice is yours."
echo "You may have to re-run this script after you install Microsoft Office updates"
