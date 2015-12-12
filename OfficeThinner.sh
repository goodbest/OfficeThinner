#!/bin/bash

# This script is only for Office 2016 for Mac.
# Tested on Office version 15.16 (151105)
# Tested on Office version 15.17 (151206)
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

#You must keep the following line, in order to deal with spaces in the FOR loop
IFS=""
appPathArray=( $ExcelPATH $PowerPointPATH $OutlookPATH $OneNotePATH )

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
echo "Thinning Fonts, it saves you ~1.4G space"
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
echo "Thinning Proofing Tools, it saves you ~1.5G space"
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
# ==============================
frameworkPATH="/Contents/Frameworks/MicrosoftOffice.framework/Versions/A/"
frameworkName="Resources"
echo "Thinning MicrosoftOffice.framework, it saves you ~0.8G space"
for appPATH in "${appPathArray[@]}";
do
    appName=${appPATH/.app/}
    mkdir -p "$backupPATH$appName$frameworkPATH"
    # echo "$backupPATH$appName$frameworkPATH"
    sudo mv "$basePATH$appPATH$frameworkPATH$frameworkName" "$backupPATH$appName$frameworkPATH"
    sudo ln -s "$basePATH$WordPATH$frameworkPATH$frameworkName" "$basePATH$appPATH$frameworkPATH$frameworkName"
    # echo "$basePATH$appPATH$frameworkPATH$frameworkName" "$backupPATH$appName$frameworkPATH"
    # echo "$basePATH$WordPATH$frameworkPATH$frameworkName" "$basePATH$appPATH$frameworkPATH$frameworkName"
done

echo ""
echo "Office Thinning Complete."
echo "The duplicate files are backed up at $backupPATH"
echo ""
echo "If everything is OK, you may delete them. But the choice is yours."
echo "You may have to re-run this script after you install Microsoft Office updates"
