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

# Configs. You can change paths here
basePath="/Applications/"
backupPath="$HOME/Desktop/OfficeThinnerBackup/"

# Symbolic link the relevant files in the apps to the first one
apps=( 'Microsoft '{Word,Excel,Powerpoint,Outlook,Onenote} )

# link_this RELATIVE_PATH_FROM_.app_DIRs...
# if you want to debug this, well.. uncomment the `set`s. Human (like me) make mistakes.
link_this(){
    # set -xv
    local officeApp
    # printf >&2 'Processing'
    # printf >&2 ' [%s]' "$@"
    # printf >&2 '...\n'
    for officeApp in "${apps[@]:1}"; do
        # printf >&2 '--> In %s...\n' "$officeApp"
        link_that "${apps[0]}" "$officeApp" "$@"
    done
    # set +xv
}
# link_that sourceApp targetApp paths_to_eat...
# Adding `-v` to `mv` and `ln` may potentially make users happier..
link_that(){
    local thing thing_bk thing_dn
    for thing; do
        # skip inexist ones.. should reduce mess.
        [[ -e $1.app/$thing/ ]] || continue
        thing_dn=$(dirname "$thing")
        thing_bk="$backupPath/$2.app/$thing_dn/"
        mkdir -p "$thing_bk"
        sudo mv "$2.app/$thing" "$thing_bk/"
        # Using the (possibly) least ambiguous POSIX ln call.. (-> dir)
        sudo ln -s "$1.app/$thing/" "$2.app/$thing_dn/"
    done
}

# Disk Usage Display
diskUsage(){
    local apps_pth
    apps_pth=("${apps[@]/%/.app}")
    apps_pth=("${apps[@]/#/${1:-$basePath}")
    du -sh "${apps_pth[@]}"
}

echo "Before running this script, Office is taking:"
diskUsage

# ==============================
# Phase I: Deal with Fonts
# Comparison Result:  Word = Excel = Powerpoint, Outlook and Onenote are subsets of Word
# ==============================
if [ -d "$basePath/${apps[0]}.app/Contents/Resources/DFonts" ]; then
    fontDir="DFonts"
else
    fontDir="Fonts"
fi
echo "Thinning Fonts, it saves you ~1.4G space"
link_this "Contents/Resources/$fontDir"

# ==============================
# Phase II: Deal with Proofing Tools
# Comparison Result:  Word = Excel = Powerpoint = Outlook, OneNote is subset of Word
# ==============================
echo "Thinning Proofing Tools, it saves you ~1.5G space"
link_this "Contents/SharedSupport/Proofing Tools"

# ==============================
# Phase III: Deal with MicrosoftOffice.framework
# Comparison Result:  Word = Excel = Powerpoint = Outlook = OneNote
# ==============================
echo "Thinning MicrosoftOffice.framework, it saves you ~0.8G space"
# Shouldn't we just deal with the '.framework' instead?
link_this "Contents/Frameworks/MicrosoftOffice.framework/Versions/A/Resources"

echo "Office Thinning Complete."
echo ""
echo "After running this script, Office is taking:"
diskUsage
echo "The duplicate files are backed up at $backupPath"
echo "If everything is OK, you may delete these files. But the choice is yours."
echo "You may have to re-run this script after you install Microsoft Office updates"
